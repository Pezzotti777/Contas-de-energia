import requests

url = 'https://script.google.com/macros/s/AKfycbxETOc0OOWxH6NJP0JhJEMUnLyKIwMbOMLNgnFqfP_mr_pMpBK3DrTMcx8kqt1QerSS/exec"
'  # Cole a URL do seu Apps Script aqui!
data = ["TESTE DIRETO", "2025-07-14", "2025-08-01", "100,00", "22", "2025-07-14", "2025-06-14", "2025-07-14", "", "TESTE", "3", "6"]
resp = requests.post(url, json=data)
print(resp.text)