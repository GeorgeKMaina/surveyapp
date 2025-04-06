# -*- coding: utf-8 -*-
"""
Created on Wed Mar  5 22:02:25 2025
@author: Wangari Kimani
"""

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.feature_extraction.text import TfidfVectorizer
from pptx import Presentation
from pptx.util import Inches
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from io import BytesIO
import re
import os
import joblib
import gdown
from bertopic import BERTopic
from dotenv import load_dotenv
import openai
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module='tqdm')
from text_utils import clean_text

#import sys
#sys.path.append(r"C:\Users\Wangari Kimani\Downloads\powerpoint app")  # 



# Load environment variables
#dotenv_path = r"C:\Users\Wangari Kimani\Downloads\powerpoint app\.env"
#load_dotenv(dotenv_path=dotenv_path)

# Verify API key
#api_key = os.getenv("OPENAI_API_KEY")
#if not api_key:
#    raise ValueError("OPENAI_API_KEY not found. Check your .env file path.")

# Initialize OpenAI client
#client = OpenAI(api_key=api_key)

# Load .env file ONLY if it exists (local environment)
if os.path.exists(".env"):
    load_dotenv()

# Try to get the API key from environment or Streamlit secrets
api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")

# Raise an error if not found
if not api_key:
    raise ValueError("OPENAI_API_KEY not found. Make sure it's in your .env file (locally) or Streamlit secrets (deployment).")

# Use the key with OpenAI
openai.api_key = api_key


# Load saved BERTopic model
# Google Drive file ID
file_id = "1vUjUBgqySAicfWYA7r7rn3RprVH3YKJy"
url = f"https://drive.google.com/uc?id={file_id}"

# File name to save locally
model_path = "best_topic_model.pkl"

# Download if not already present
if not os.path.exists(model_path):
    gdown.download(url, model_path, quiet=False)

# Load the BERT topic model
bert_model = joblib.load(model_path)
topic_names = joblib.load("topic_label_map.pkl")

# Load trained question type classifier
#classifier_model = joblib.load(r"C:\Users\Wangari Kimani\Downloads\powerpoint app\best_model.pkl")
classifier_model = joblib.load("best_model.pkl")
# Streamlit UI
st.title("Survey Data Analysis & PowerPoint Report Generator")

survey_type_options = ["Employee Feedback", "Customer Satisfaction", "Market Research", "Other"]
survey_type = st.selectbox("Select Survey Type", survey_type_options)
if survey_type == "Other":
    survey_type = st.text_input("Describe the survey type")

data_collection_date = st.date_input("Date of Data Collection")

uploaded_file = st.file_uploader("Upload Survey Dataset (Excel)", type=["xlsx"])

if uploaded_file:
    df = pd.ExcelFile(uploaded_file)
    sheet_name = df.sheet_names[0]
    data = df.parse(sheet_name)

    st.write("Dataset Preview:")
    st.dataframe(data.head())

    # --- Use trained model to classify questions ---
    # Use the same clean_text function used in training
    question_texts = list(data.columns)
    cleaned_questions = [clean_text(q) for q in question_texts]
    predicted_labels = classifier_model.predict(cleaned_questions)

# Normalize predictions to match expected format
    normalized_labels = []
    for label in predicted_labels:
        l = label.lower().replace(" ", "-")
        if l.startswith("open"):
            normalized_labels.append("open-ended")
        else:
            normalized_labels.append("closed-ended")

    
    question_types = dict(zip(question_texts, normalized_labels))
    open_ended = [q for q, label in question_types.items() if label == 'open-ended']
    closed_ended = [q for q, label in question_types.items() if label == 'closed-ended']


    # Preprocess text for open-ended
    def preprocess_text(text):
        if isinstance(text, str):
            text = text.lower()
            text = re.sub(r'[^a-zA-Z\s]', '', text)
            return text
        return ""

    data_cleaned = data[open_ended].astype(str).applymap(preprocess_text)

    # Analyze open-ended responses with BERTopic
    topic_freqs = {}  # {question: {topic: frequency}}

    for question in open_ended:
        # Clean responses to keep only proper text
        responses = data_cleaned[question].dropna().tolist()
        responses = [r for r in responses if isinstance(r, str) and len(r.strip()) > 3]
        
        if len(responses) > 5:
            try:
                topics, _ = bert_model.fit_transform(responses)
                topic_counts = pd.Series(topics).value_counts().sort_values(ascending=False)
                topic_freqs[question] = {
                    bert_model.get_topic(topic)[0][0]: count
                    for topic, count in topic_counts.items() if topic != -1
        }
                
            except Exception as e:
                topic_freqs[question] = {f"BERTopic Error: {str(e)}": 1}
        else:
            topic_freqs[question] = {"Not enough responses": 1}



    # Create PowerPoint
    prs = Presentation()

    # Executive Summary
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Executive Summary"
    textbox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
    text_frame = textbox.text_frame
    summary_text = (
        f"Survey Type: {survey_type}\n"
        f"Data Collection Date: {data_collection_date.strftime('%Y-%m-%d')}\n"
        f"Total Questions: {len(data.columns)}\n"
        f"Open-Ended Questions: {len(open_ended)}\n"
        f"Closed-Ended Questions: {len(closed_ended)}"
    )
    text_frame.text = summary_text

    # Closed-ended analysis
    for question in closed_ended:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = question
        responses = data[question].value_counts(normalize=True) * 100

        if responses.empty:
            responses = pd.Series([0], index=['No Data'])

        chart_data = CategoryChartData()
        chart_data.categories = list(responses.index)
        chart_data.add_series("Responses", tuple(responses.values))

        slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(2), Inches(8), Inches(4), chart_data
        )

        prompt = f"""
        You are a data analyst preparing a strategic summary for a survey on {survey_type}.
        
        Survey Question: {question}
        Responses: {responses.to_dict()}
        
        Analyze the results by doing the following:
        1. Interpret what the distribution suggests about respondent attitudes or behavior.
        2. Suggest actionable decisions or follow-ups based on this.
        3. Flag any biases, assumptions, or interpretation limitations.
        
        Write 3–5 numbered bullet points. Be clear, concise, and strategic.
        """
        
        #prompt = f"Generate insights from the following survey data distribution: {responses.to_dict()}"
        try:
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}]
            )
            narrative = response.choices[0].message.content
        except Exception as e:
            narrative = f"Error generating narrative: {e}"

        textbox = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(8), Inches(2))
        text_frame = textbox.text_frame
        text_frame.text = narrative

    # Open-ended analysis
    #for question, topics in topic_freqs.items():
    #    slide = prs.slides.add_slide(prs.slide_layouts[5])
    #    slide.shapes.title.text = question

    #    chart_data = CategoryChartData()
     #   chart_data.categories = list(topics.keys())
      #  chart_data.add_series("Frequency", list(topics.values()))

       # slide.shapes.add_chart(
        #    XL_CHART_TYPE.BAR_CLUSTERED, Inches(1), Inches(2), Inches(8), Inches(4), chart_data
        #)

#        textbox = slide.shapes.add_textbox(Inches(1), Inches(5.2), Inches(8), Inches(1.5))
   #     text_frame = textbox.text_frame
    #    text_frame.text = "Most discussed topics from responses."
    if topics and isinstance(topics, dict) and len(topics.keys()) > 0:
        categories = list(topics.keys())
        values = list(topics.values())
    
        # Check that categories are valid strings
        if all(isinstance(cat, str) for cat in categories) and len(categories) > 0:
            chart_data = CategoryChartData()
            chart_data.categories = categories
            chart_data.add_series("Frequency", values)
    
            slide.shapes.add_chart(
                XL_CHART_TYPE.BAR_CLUSTERED, Inches(1), Inches(2), Inches(8), Inches(4), chart_data
            )
    
            textbox = slide.shapes.add_textbox(Inches(1), Inches(5.2), Inches(8), Inches(1.5))
            text_frame = textbox.text_frame
            text_frame.text = "Most discussed topics from responses."
        else:
            # No valid categories
            textbox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
            text_frame = textbox.text_frame
            text_frame.text = "No valid topics found for this question."
    else:
        textbox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
        text_frame = textbox.text_frame
        text_frame.text = "No topics generated or BERTopic failed."

    # Export PPT
    ppt_stream = BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)

    st.success("Analysis complete! Download your PowerPoint report below.")
    st.download_button(
        "Download PowerPoint Report",
        ppt_stream,
        "survey_report.pptx",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
