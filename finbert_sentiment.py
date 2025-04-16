# -*- coding: utf-8 -*-
"""FinBert Sentiment.ipynb

Automatically generated by Colab.

Original file is located at
    https://colab.research.google.com/drive/17_Ui7FQc0qeITwI_OAqQvBx_A8ic-Mhw
"""

pip install transformers torch pandas openpyxl

import pandas as pd
import torch
from transformers import AutoTokenizer, AutoModelForSequenceClassification
from torch.nn.functional import softmax
from tqdm import tqdm

# Load FinBERT model + tokenizer
model_name = "yiyanghkust/finbert-tone"
tokenizer = AutoTokenizer.from_pretrained(model_name)
model = AutoModelForSequenceClassification.from_pretrained(model_name)

# Use GPU if available
device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
model.to(device)
model.eval()

# Load your Excel/CSV file
df = pd.read_excel("CRYPTO_news.xlsx")  # ← replace with your filename

# Function to preprocess in batches
def get_sentiment_batch(texts, batch_size=64):
    sentiments = []
    confidences = []

    for i in tqdm(range(0, len(texts), batch_size)):
        batch_texts = list(texts[i:i+batch_size])
        inputs = tokenizer(batch_texts, padding=True, truncation=True, return_tensors="pt").to(device)

        with torch.no_grad():
            outputs = model(**inputs)
            probs = softmax(outputs.logits, dim=1)
            max_probs, predictions = torch.max(probs, dim=1)

        labels = [model.config.id2label[pred.item()] for pred in predictions]
        sentiments.extend(labels)
        confidences.extend([round(prob.item(), 4) for prob in max_probs])

    return sentiments, confidences

# Run sentiment analysis
df['TITLE'] = df['TITLE'].astype(str)
sentiments, confidences = get_sentiment_batch(df['TITLE'])

# Store results
df['Sentiment'] = sentiments
df['Confidence'] = confidences

# Save to Excel
df.to_excel("crypto_news_with_sentiment.xlsx", index=False)
print("✅ Sentiment analysis completed and saved.")

