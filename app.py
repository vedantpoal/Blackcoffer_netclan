import pandas as pd
import requests
from newspaper import Article
import os
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
import re
import syllables

import nltk
nltk.download('punkt_tab')

nltk.download('punkt', quiet=True)  

# Load input data
input_df = pd.read_excel('D:/projects/new assignment/Input.xlsx')

# Load stop words
def load_stop_words(filenames):
    stop_words = set()
    for filename in filenames:
        try:
            with open(filename, 'r', encoding='utf-8', errors='ignore') as file:
                content = file.read().splitlines()
                for line in content:
                    parts = line.split('|')
                    stop_words.update(part.strip().upper() for part in parts[0].split())
        except UnicodeDecodeError:
            with open(filename, 'r', encoding='ISO-8859-1', errors='ignore') as file:
                content = file.read().splitlines()
                for line in content:
                    parts = line.split('|')
                    stop_words.update(part.strip().upper() for part in parts[0].split())
    return stop_words

stop_words_files = [
    'D:/projects/new assignment/StopWords/StopWords_Auditor.txt',
    'D:/projects/new assignment/StopWords/StopWords_Currencies.txt',
    'D:/projects/new assignment/StopWords/StopWords_DatesandNumbers.txt',
    'D:/projects/new assignment/StopWords/StopWords_Generic.txt',
    'D:/projects/new assignment/StopWords/StopWords_GenericLong.txt',
    'D:/projects/new assignment/StopWords/StopWords_Geographic.txt',
    'D:/projects/new assignment/StopWords/StopWords_Names.txt'
]
stop_words = load_stop_words(stop_words_files)

# Load positive and negative words
try:
    with open('D:/projects/new assignment/MasterDictionary/positive-words.txt', 'r', encoding='utf-8') as file:
        positive_words = set(file.read().splitlines())
except UnicodeDecodeError:
    with open('D:/projects/new assignment/MasterDictionary/positive-words.txt', 'r', encoding='ISO-8859-1') as file:
        positive_words = set(file.read().splitlines())

try:
    with open('D:/projects/new assignment/MasterDictionary/negative-words.txt', 'r', encoding='utf-8') as file:
        negative_words = set(file.read().splitlines())
except UnicodeDecodeError:
    with open('D:/projects/new assignment/MasterDictionary/negative-words.txt', 'r', encoding='ISO-8859-1') as file:
        negative_words = set(file.read().splitlines())

# Text cleaning function
def clean_text(text):
    text = re.sub(r'[^a-zA-Z\s]', '', text)  
    words = word_tokenize(text.upper())
    filtered_words = [word for word in words if word not in stop_words and len(word) > 1]
    return filtered_words

# Syllable count with exceptions
def count_syllables(word):
    word = word.lower()
    if word.endswith(('es', 'ed')) and len(word) > 3:
        return max(1, syllables.estimate(word) - 1)
    return syllables.estimate(word)

# Personal pronouns regex
personal_pronouns = re.compile(r'\b(I|we|my|ours|us)\b', re.IGNORECASE)

# Analyze article text
def analyze_text(text):
    sentences = sent_tokenize(text)
    words = clean_text(text)
    cleaned_text = ' '.join(words).lower()
    
    # Sentiment scores
    positive_score = sum(1 for word in words if word.lower() in positive_words)
    negative_score = sum(1 for word in words if word.lower() in negative_words)
    polarity = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity = (positive_score + negative_score) / (len(words) + 0.000001)
    
    # Readability
    avg_sentence_length = len(words) / len(sentences) if sentences else 0
    complex_words = [word for word in words if count_syllables(word) > 2]
    complex_word_count = len(complex_words)
    percentage_complex = (complex_word_count / len(words)) * 100 if words else 0
    fog_index = 0.4 * (avg_sentence_length + percentage_complex)
    
    # Word and syllable stats
    word_count = len(words)
    syllable_per_word = sum(count_syllables(word) for word in words) / word_count if word_count else 0
    avg_word_length = sum(len(word) for word in words) / word_count if word_count else 0
    
    # Personal pronouns (excluding 'US' country)
    pronouns = len(re.findall(personal_pronouns, text))
    
    return {
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': negative_score,
        'POLARITY SCORE': polarity,
        'SUBJECTIVITY SCORE': subjectivity,
        'AVG SENTENCE LENGTH': avg_sentence_length,
        'PERCENTAGE OF COMPLEX WORDS': percentage_complex,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_sentence_length,
        'COMPLEX WORD COUNT': complex_word_count,
        'WORD COUNT': word_count,
        'SYLLABLE PER WORD': syllable_per_word,
        'PERSONAL PRONOUNS': pronouns,
        'AVG WORD LENGTH': avg_word_length
    }

# Process each URL
output_data = []
output_dir = 'D:/projects/new assignment/data'
os.makedirs(output_dir, exist_ok=True)

for index, row in input_df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    
    try:
        # Extract article
        article = Article(url)
        article.download()
        article.parse()
        
        # Save text
        with open(os.path.join(output_dir, f'{url_id}.txt'), 'w', encoding='utf-8') as file:
            file.write(f"{article.title}\n\n{article.text}")
        
        # Analyze
        analysis = analyze_text(article.text)
        analysis['URL_ID'] = url_id
        analysis['URL'] = url
        output_data.append(analysis)
        
    except Exception as e:
        print(f"Error processing {url_id}: {e}")
        output_data.append({
            'URL_ID': url_id,
            'URL': url,
            **{key: None for key in analyze_text("").keys()}
        })

# Create output DataFrame
output_df = pd.DataFrame(output_data)
columns = ['URL_ID', 'URL'] + list(analyze_text("").keys())
output_df = output_df[columns]

# Save to Excel
output_df.to_excel('D:/projects/new assignment/Output Data Structure.xlsx', index=False)