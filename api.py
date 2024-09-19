import requests

url = 'http://127.0.0.1:5000/upload'
file_path = r'C:\Users\klopmans\Documents\רחלי יוסקוביץ\FullStack developer.docx'

with open(file_path, 'rb') as file:
    files = {'file': (file_path, file, '')}
    response = requests.post(url, files=files)

print(response.text)