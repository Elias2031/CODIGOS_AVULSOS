import requests
import json

host = "www2.agencianet.fazenda.df.gov.br"
url = f"https://{host}/DEC/ContarDocumentosDECR"

headers = {
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json;charset=UTF-8",
    "Origin": f"https://{host}",
    "Referer": f"https://{host}/DEC/",
    "User-Agent": "Mozilla/5.0"
}

cookies = {
    # Se for via DOMÍNIO:
    "CFID": "...",
    "CFTOKEN": "...",
    "cf_clearance": "...",
    "ASP.NET_SessionId": "..."
}

payload = {
    "Origem": "E",
    "Tipo": "NFE",
    "Criterio": "D",
    "CpfCnpj": "04550020000197",
    "DataInicio": "2024-01-01T03:00:00.000Z",
    "DataFim": "2024-01-31T03:00:00.000Z",
    "Offset": 0,
    "Count": 10001,
    "captcha": "SEU_TOKEN_VALIDO_AQUI"
}

r = requests.post(url, headers=headers, cookies=cookies, json=payload, timeout=20)
print(r.status_code)
print(r.text[:1000])
