# -*- coding: utf-8 -*-
"""
Created on Sun Mar 23 18:39:46 2025

@author: Wangari Kimani
"""

import re

def clean_text(text: str) -> str:
    text = re.sub(r'^\d+\.\s*', '', text)
    text = text.lower()
    text = re.sub(r'[^\w\s()]+', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text
