import requests
import os, json
# 從環境變數讀取 API_URL 和 Headers
API_URL = os.getenv('API_URL')
HEADERS_JSON = os.getenv('HEADERS_JSON')
Headers = json.loads(HEADERS_JSON)

def call_prediction_api(question):
    payload = {"question": question}
    try:
        print('qa')
        response = requests.post(API_URL, headers=Headers, json=payload)
        response.raise_for_status()
        return response.json().get("text", "No answer available.")
    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}")
        return f"HTTP error occurred: {http_err}"
    except Exception as err:
        print(f"Other error occurred: {err}")
        return f"Other error occurred: {err}"
