import os
import pandas as pd
import requests

# Configuração da pasta local (simulando o SFTP)
pasta_local = "C:\\sftp_simulado"

# Configuração da API Lemontech
URL_API = "https://treinamento.lemontech.com.br:443/wsselfbooking/WsSelfBookingService"  # Substitua pela URL real da API
HEADERS = {"Content-Type": "text/xml"}
USER = "6d64cbcadc5981ef1387fb4a7a63e7ae"
PASSWORD = "87333acf7b863397097848232f275803"
KEY_CLIENT = "integracao_homolog_qa"

# Procurar o arquivo Excel na pasta
arquivos = os.listdir(pasta_local)
arquivo_excel = None
for arquivo in arquivos:
    if arquivo.endswith(".xls") or arquivo.endswith(".xlsx"):
        arquivo_excel = os.path.join(pasta_local, arquivo)
        break

if not arquivo_excel:
    print("⚠️ Nenhum arquivo Excel encontrado na pasta.")
    exit()

print(f"📄 Arquivo encontrado: {arquivo_excel}")

# Ler o arquivo Excel
df = pd.read_excel(arquivo_excel)

# Selecionar as colunas necessárias
colunas_desejadas = [
    "codigoCentroDeCusto", 
    "descricaoCentroDeCusto", 
    "codigoRegional", 
    "autoAprovavel", 
    "debitaBudget"
]

# Verificar se todas as colunas existem
colunas_faltando = [col for col in colunas_desejadas if col not in df.columns]
if colunas_faltando:
    print(f"⚠️ As seguintes colunas não foram encontradas: {colunas_faltando}")
    exit()

# Converter os valores das colunas booleanas ("S" -> True, "N" -> False)
df["autoAprovavel"] = df["autoAprovavel"].map(lambda x: True if x == "S" else False)
df["debitaBudget"] = df["debitaBudget"].map(lambda x: True if x == "S" else False)

# Iterar por cada linha do arquivo e enviar a requisição para a API
for idx, row in df.iterrows():
    codigo = row["codigoCentroDeCusto"]
    descricao = row["descricaoCentroDeCusto"]
    regional = row["codigoRegional"]
    auto_aprovavel = str(row["autoAprovavel"]).lower()  # Converte para 'true' ou 'false'
    debita_budget = str(row["debitaBudget"]).lower()

    # Construir o XML SOAP conforme a documentação
    xml_request = f"""<?xml version="1.0" encoding="UTF-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
xmlns:ser="http://lemontech.com.br/selfbooking/wsselfbooking/services">
    <soapenv:Header>
        <ser:userPassword>{PASSWORD}</ser:userPassword>
        <ser:userName>{USER}</ser:userName>
        <ser:keyClient>{KEY_CLIENT}</ser:keyClient>
    </soapenv:Header>
    <soapenv:Body>
        <ser:cadastrarCentroDeCusto>
            <centroDeCusto>
                <codigo>{codigo}</codigo>
                <regionalRef>
                    <codigo>{regional}</codigo>
                </regionalRef>
                <descricao>{descricao}</descricao>
                <configuracao>
                    <autoAprovavel>{auto_aprovavel}</autoAprovavel>
                    <debitaBudget>{debita_budget}</debitaBudget>
                </configuracao>
                <ativo>true</ativo>
            </centroDeCusto>
        </ser:cadastrarCentroDeCusto>
    </soapenv:Body>
</soapenv:Envelope>"""

    # Enviar a requisição para a API
    response = requests.post(URL_API, data=xml_request, headers=HEADERS)

    print(f"\n📡 Enviando centro de custo: {codigo} - {descricao}")
    print(f"🔄 Status: {response.status_code}")
    print("📬 Resposta da API:")
    print(response.text)

print("\n✅ Processo concluído!")
