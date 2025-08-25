import requests

# Substitua pela sua URL do Apps Script (web app)
url = "https://script.google.com/macros/s/AKfycbxmuHslwyjlRcwc3V432oBvW-9LzxRQkddcYbx2Bl1Iz-GiZDYn5oNysqxJzKiG8axU/exec"

# Exemplo de dados (ajuste a quantidade e ordem para bater com suas colunas!)
data = [
    "Teste Instalação", "2025-07-14", "2025-08-01", "100,00", "22", "2025-07-14", "2025-06-14",
    "2025-07-14", "", "TESTE", "3", "6"
]

resp = requests.post(url, json=data)
print(resp.text)
