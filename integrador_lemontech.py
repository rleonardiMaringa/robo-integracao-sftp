import os
import pandas as pd
import requests

# === Configuração ===
PASTA_LOCAL = "C:\\sftp_simulado"

URL_API = "https://treinamento.lemontech.com.br:443/wsselfbooking/WsSelfBookingService"
HEADERS = {"Content-Type": "text/xml"}
USER = "6d64cbcadc5981ef1387fb4a7a63e7ae"
PASSWORD = "87333acf7b863397097848232f275803"
KEY_CLIENT = "integracao_homolog_qa"

# === Enviar centro de custo ===
def enviar_centro_de_custo():
    print("📤 Enviando centros de custo...")

    for arquivo in os.listdir(PASTA_LOCAL):
        if "centro" in arquivo.lower() and arquivo.endswith((".xls", ".xlsx")):
            df = pd.read_excel(os.path.join(PASTA_LOCAL, arquivo))
            break
    else:
        print("❌ Nenhum arquivo de centro de custo encontrado.")
        return

    colunas = ["codigoCentroDeCusto", "descricaoCentroDeCusto", "codigoRegional", "autoAprovavel", "debitaBudget"]
    if not all(col in df.columns for col in colunas):
        print("❌ Colunas obrigatórias faltando no centro de custo.")
        return

    df["autoAprovavel"] = df["autoAprovavel"].map(lambda x: str(x).strip().upper() == "S")
    df["debitaBudget"] = df["debitaBudget"].map(lambda x: str(x).strip().upper() == "S")

    for _, row in df.iterrows():
        xml = f"""<?xml version="1.0" encoding="UTF-8"?>
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
                <codigo>{row['codigoCentroDeCusto']}</codigo>
                <regionalRef>
                    <codigo>{row['codigoRegional']}</codigo>
                </regionalRef>
                <descricao>{row['descricaoCentroDeCusto']}</descricao>
                <configuracao>
                    <autoAprovavel>{str(row['autoAprovavel']).lower()}</autoAprovavel>
                    <debitaBudget>{str(row['debitaBudget']).lower()}</debitaBudget>
                </configuracao>
                <ativo>true</ativo>
            </centroDeCusto>
        </ser:cadastrarCentroDeCusto>
    </soapenv:Body>
</soapenv:Envelope>"""

        response = requests.post(URL_API, data=xml, headers=HEADERS)
        print(f"✅ Enviado: {row['codigoCentroDeCusto']} - {row['descricaoCentroDeCusto']} | Status: {response.status_code}")
        print("📨 Resposta da API:")
        print(response.text)

# === Enviar funcionário ===
def enviar_funcionarios():
    print("📤 Enviando funcionários...")

    for arquivo in os.listdir(PASTA_LOCAL):
        if "funcionario" in arquivo.lower() and arquivo.endswith((".xls", ".xlsx")):
            df = pd.read_excel(os.path.join(PASTA_LOCAL, arquivo))
            break
    else:
        print("❌ Nenhum arquivo de funcionário encontrado.")
        return

    # De-paras
    mapa_genero = {'F': 'FEMININO', 'M': 'MASCULINO'}
    mapa_perfil_funcionario = {'2': 'GESTOR', '3': 'APROVADOR', '4': 'SOLICITANTE', '6': 'PASSAGEIRO', '10': 'APROVADOR MASTER'}
    mapa_perfil_aereo = {'1': 'ECONOMICA', '2': 'EXECUTIVA', '3': 'PRIMEIRA_CLASSE', '4': 'ECONOMICA_PLUS'}

    for _, row in df.iterrows():
        if str(row['operacao']).strip().upper() != "A":
            continue

        # Transformações
        data_nascimento = pd.to_datetime(row['dataNascimento'], dayfirst=True).strftime('%Y-%m-%d') if pd.notnull(row['dataNascimento']) else ''
        sexo = mapa_genero.get(str(row['genero']).strip().upper(), '')
        perfil_func = mapa_perfil_funcionario.get(str(row['perfilFuncionario']).strip(), '')
        perfil_aereo = mapa_perfil_aereo.get(str(row['perfilAereo']).strip(), '')
        auto_aprova = str(row['autoAprovar']).strip().upper() == "S"
        solicita_para_todos = str(row['solicitaParaTodos']).strip().upper() == "S"
        bypass_nacional = str(row['aprovacaoAutomaticaNacional']).strip().upper() == "S"
        bypass_internacional = str(row['aprovacaoAutomaticaInternacional']).strip().upper() == "S"
        utiliza_logado = str(row['utilizaUsuarioLogado']).strip().upper() == "S"
        bloqueia_viajar = str(row['bloqueiaUsuarioParaViajar']).strip().upper() == "S"
        ativo = "false" if str(row['ativo']).strip().upper() == "N" else "true"

        xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ser="http://lemontech.com.br/selfbooking/wsselfbooking/services">
   <soapenv:Header>
      <ser:userPassword>{PASSWORD}</ser:userPassword>
      <ser:userName>{USER}</ser:userName>
      <ser:keyClient>{KEY_CLIENT}</ser:keyClient>
   </soapenv:Header>
   <soapenv:Body>
      <ser:cadastrarFuncionario>
         <funcionario>
            <matricula>{row['matricula']}</matricula>
            <nome>{row['nome']}</nome>
            <departamento>{row['departamento']}</departamento>
            <cargo>{row['cargo']}</cargo>
            <cpf>{row['cpf']}</cpf>
            <dataNascimento>{data_nascimento}</dataNascimento>
            <sexo>{sexo}</sexo>
            <subCentroDeCustoRef>
               <codigo>{row['codigoSubCentroDeCusto']}</codigo>
               <centroDeCustoRef>
                  <codigo>{row['codigoCentroDeCusto']}</codigo>
                  <regionalRef>
                     <codigo>{row['codigoRegional']}</codigo>
                  </regionalRef>
               </centroDeCustoRef>
            </subCentroDeCustoRef>
            <contato>
               <email>{row['email']}</email>
               <ddiTelefone>{row['ddiTelefone']}</ddiTelefone>
               <dddTelefone>{row['dddTelefone']}</dddTelefone>
               <telefone>{row['telefone']}</telefone>
               <ddiCelular>{row['ddiCelular']}</ddiCelular>
               <dddCelular>{row['dddCelular']}</dddCelular>
               <celular>{row['celular']}</celular>
               <forcaAtualizacao>true</forcaAtualizacao>
            </contato>
            <login>
               <usuario>{row['login']}</usuario>
            </login>
            <bypassAprovacaoNacional>{str(bypass_nacional).lower()}</bypassAprovacaoNacional>
            <bypassAprovacaoInternacional>{str(bypass_internacional).lower()}</bypassAprovacaoInternacional>
            <configuracao>
               <autoAprova>{str(auto_aprova).lower()}</autoAprova>
               <solicitaParaTodos>{str(solicita_para_todos).lower()}</solicitaParaTodos>
               <categoriaHospedagem>{row['categoriaHospedagem']}</categoriaHospedagem>
               <perfilFuncionario>{perfil_func}</perfilFuncionario>
               <perfilAereo>{perfil_aereo}</perfilAereo>
               <utilizaUsuarioLogado>{str(utiliza_logado).lower()}</utilizaUsuarioLogado>
               <bloqueiaUsuarioParaViajar>{str(bloqueia_viajar).lower()}</bloqueiaUsuarioParaViajar>
               <emailEnvioCopiaDeVoucher>{row.get('emailEnvioCopiaDeVoucher', '')}</emailEnvioCopiaDeVoucher>
            </configuracao>
            <ativo>{ativo}</ativo>
         </funcionario>
      </ser:cadastrarFuncionario>
   </soapenv:Body>
</soapenv:Envelope>"""

        response = requests.post(URL_API, data=xml, headers=HEADERS)
        print(f"✅ Enviado funcionário: {row['matricula']} - {row['nome']} | Status: {response.status_code}")
        print("📨 Resposta da API:")
        print(response.text)


# === Execução principal ===
if __name__ == "__main__":
    enviar_centro_de_custo()
    enviar_funcionarios()
    print("\n🏁 Finalizado!")
