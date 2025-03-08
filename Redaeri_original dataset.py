import requests
import pandas as pd
import json

# API 핸들러 클래스
class CompletionExecutor:
    def __init__(self, host, api_key, request_id):
        self._host = host
        self._api_key = api_key
        self._request_id = request_id

    def execute(self, completion_request):
        headers = {
            'Authorization': f'Bearer {self._api_key}',   # Bearer 인증 방식
            'X-NCP-CLOVASTUDIO-REQUEST-ID': self._request_id,
            'Content-Type': 'application/json; charset=utf-8',
            'Accept': 'application/json'
        }
        response = requests.post(self._host + '/testapp/v1/chat-completions/HCX-003',
                                 headers=headers, json=completion_request)
        if response.status_code == 200:
            return response.json()['result']['message']['content']
        else:
            return 'Error: ' + response.text

# 엑셀 처리 함수
def process_excel(input_file, output_file, completion_executor):
    df = pd.read_excel(input_file)
    
    if 'completion' not in df.columns:
        df['completion'] = ''
    else:
        df['completion'] = df['completion'].astype(str)

    for index, row in df.iterrows():
        system_content = row['system'] if 'system' in df.columns else ""
        user_content = row['user'] if 'user' in df.columns else ""
        
        preset_text = [
            {"role": "system", "content": system_content},
            {"role": "user", "content": user_content}
        ]
        
        request_data = {
            'messages': preset_text,
            'topP': 0.8,
            'topK': 0,
            'maxTokens': 256,
            'temperature': 0.8,
            'repeatPenalty': 5.0,
            'stopBefore': [],
            'includeAiFilters': True,
            'seed': 0
        }
        
        response = completion_executor.execute(request_data)
        df.at[index, 'completion'] = response

    df.to_excel(output_file, index=False)

# 메인 실행 블록
if __name__ == '__main__':
    # API 인증 정보 설정
    API_KEY = '{key}'  # 클로바 Studio에서 발급받은 최신 API 키
    REQUEST_ID = 'test01'
    INPUT_FILENAME = "/Users/jjshim/Desktop/성능평가/RDR_01.xlsx"
    OUTPUT_FILENAME = "/Users/jjshim/Desktop/성능평가/RDR_01_result.xlsx"

    completion_executor = CompletionExecutor(
        host='https://clovastudio.stream.ntruss.com',
        api_key=API_KEY,
        request_id=REQUEST_ID
    )
    process_excel(INPUT_FILENAME, OUTPUT_FILENAME, completion_executor)
