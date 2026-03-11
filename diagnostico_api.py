import requests, sys, json

HOST = "www2.agencianet.fazenda.df.gov.br"
IP_DEBUG = "131.72.223.132"  # ou None para não forçar
url = f"https://{HOST}/DEC/ContarDocumentosDECR"
headers = {
    "Content-Type": "application/json;charset=UTF-8",
    "Origin": f"https://{HOST}",
    "Referer": f"https://{HOST}/DEC/",
    "X-Requested-With": "XMLHttpRequest",
    "User-Agent": "morpho-agent/1.0"
}

cookies = {
    # opcional: preencha se você tem sessão do navegador
    # "ASP.NET_SessionId": "...",
    # "cf_clearance": "..."
}

payload = {
    "Origem":"E","Tipo":"NFE","Criterio":"D",
    "CpfCnpj":"04550020000197",
    "DataInicio":"2024-01-01T03:00:00.000Z",
    "DataFim":"2024-01-31T03:00:00.000Z",
    "Offset":0,"Count":10001,"captcha":"TOKEN_VÁLIDO_AQUI"
}

s = requests.Session()
s.headers.update(headers)
s.cookies.update(cookies)

if IP_DEBUG:
    # não existe --resolve nativo no requests; faz a requisição ao IP e manda Host header
    url_ip = f"https://{IP_DEBUG}/DEC/ContarDocumentosDECR"
    resp = s.post(url_ip, json=payload, headers={"Host": HOST}, verify=True, timeout=20)
else:
    resp = s.post(url, json=payload, timeout=20)

print("Status:", resp.status_code)
print("Content-Type:", resp.headers.get("Content-Type"))
print("Redirected:", resp.is_redirect or resp.history)
if resp.history:
    for r in resp.history:
        print("  ->", r.status_code, r.headers.get("Location"))
print("Response snippet:")
print(resp.text[:1000])
