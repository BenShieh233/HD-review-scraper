import streamlit as st
import requests
import pandas as pd
import json
import re
import time
import io
from bs4 import BeautifulSoup

def clean_client_response(text):
    if text:
        soup = BeautifulSoup(text, 'html.parser')
        cleaned_text = soup.get_text(separator=" ").strip()
        cleaned_text = re.sub(r'\s+', ' ', cleaned_text)
    else:
        cleaned_text = None
    return cleaned_text

def extract_review(review):
    BadgesOrder = review.get('BadgesOrder', [])
    BadgesOrder = BadgesOrder[0] if BadgesOrder else None

    Age = review.get('ContextDataValues', {}).get('Age', None)
    if Age:
        Age = Age.get('Value', None)
    
    IsRecommended = review.get('IsRecommended', None)
    IsSyndicated = review.get('IsSyndicated', None)
    ProductId = review.get('ProductId', None)
    SubmissionTime = review.get('SubmissionTime', None)
    Rating = review.get('Rating', None)
    Title = review.get('Title', None)
    ReviewText = review.get('ReviewText', None)
    
    ClientResponses = review.get('ClientResponses', [])
    ClientResponses_Text = ClientResponses[0].get('Response') if ClientResponses else None
    ClientResponses_Date = ClientResponses[0].get('Date') if ClientResponses else None
    ClientResponses_Department = ClientResponses[0].get('Department') if ClientResponses else None

    review_data = {
        'BadgesOrder': BadgesOrder,
        'Age': Age,
        'IsRecommended': IsRecommended,
        'IsSyndicated': IsSyndicated,
        'ProductId': ProductId,
        'SubmissionTime': SubmissionTime,
        'Rating': Rating,
        'Title': Title,
        'ReviewText': ReviewText,
        'ClientResponses': clean_client_response(ClientResponses_Text),
        'ClientResponses_Date': ClientResponses_Date,
        'ClientResponses_Department': ClientResponses_Department
    }
    return review_data

def update_payload(payload, increment):
    payload['variables']['startIndex'] = increment
    return payload

def extract_bad_reviews(url, page_num, headers, payload, selected_star_ratings):
    itemId = re.search(r'/(\d+)$', url).group(1)
    current_url = re.findall(r'/p/(.+)', url)[0]
    if not itemId:
        st.error("Invalid URL format. Please check the URL.")
        return pd.DataFrame()
    
    headers['X-Current-Url'] = current_url
    payload['variables']['itemId'] = itemId
    payload['variables']['filters']['starRatings'] = selected_star_ratings
    
    all_reviews = []
    increment = 1
    for _ in range(page_num):
        payload = update_payload(payload, increment)
        response = requests.post(
            'https://apionline.homedepot.com/federation-gateway/graphql?opname=reviews',
            json=payload, headers=headers
        )
        if response.status_code != 200:
            st.error("Failed to fetch data from Home Depot API.")
            return pd.DataFrame()
        
        data = response.json().get('data', {}).get('reviews', {}).get('Results', [])
        for review in data:
            all_reviews.append(extract_review(review))
        increment += 10
        time.sleep(3)  # 避免请求过于频繁
    
    return pd.DataFrame(all_reviews)

# Streamlit UI
st.title("Home Depot 评论抓取框架（测试版）")
url = st.text_input("请输入产品网页链接:")
page_num = st.number_input("页数（请不要超过3）:", min_value=1, max_value=10, value=2, step=1)

# **新增** 允许用户选择星级评论
star_choices = st.multiselect("选择要爬取的评论星级（可多选）:", [1, 2, 3, 4, 5], default=[1, 2, 3, 4, 5])

# 新增文件名输入
file_name_input = st.text_input("请输入保存的 Excel 文件名称（例如：reviews.xlsx）:")

if st.button("抓取评论"):
    try:
        with open("review_scraper/review_headers.json", 'r') as f:
            headers = json.load(f)
        with open("review_scraper/review_payload.json", 'r') as f:
            payload = json.load(f)

        # **转换用户选择的星级排序**
        selected_star_ratings = sorted(star_choices, reverse=True)
        print(selected_star_ratings)
        df = extract_bad_reviews(url, page_num, headers, payload, selected_star_ratings)
        if not df.empty:
            st.write(df)
            
            # 检查并自动补全文件名后缀
            file_name = file_name_input.strip() or "reviews.xlsx"
            if not file_name.lower().endswith(".xlsx"):
                file_name += ".xlsx"
            
            # 导出 Excel 文件，使用 openpyxl 引擎
            excel_buffer = io.BytesIO()
            df.to_excel(excel_buffer, index=False, engine='openpyxl')
            excel_buffer.seek(0)
            
            st.download_button(
                label="导出 Excel 表格文件",
                data=excel_buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No reviews found.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
