import os
import subprocess
import json
import time
import base64
import threading
import sys

import pandas as pd
import requests
from signxml import XMLSigner, methods
from lxml import etree
from requests_pkcs12 import post
from xml.etree import ElementTree as ET

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from pathlib import Path
from dotenv import load_dotenv

# 1) Define o diretório base do projeto (um nível acima de src/)
BASE_DIR = Path(__file__).resolve().parent.parent

# 2) Localiza e carrega o .env na raiz
env_path = BASE_DIR / '.env'
print(f"DEBUG: procurando .env em → {env_path}")
if not env_path.exists():
    print("⚠️ .env NÃO encontrado nesse caminho. Ajuste o env_path")
else:
    print("✅ .env encontrado.")
load_dotenv(dotenv_path=str(env_path))

# 3) Define o diretório onde ficam os certificados
CERTS_DIR = BASE_DIR / 'certs'

# 4) Carrega do .env os nomes dos arquivos
PFX_FILE           = os.getenv('PFX_FILE')
PFX_PASSWORD       = os.getenv('PFX_PASSWORD')
CONSUMER_KEY       = os.getenv('CONSUMER_KEY')
CONSUMER_SECRET    = os.getenv('CONSUMER_SECRET')
CONTRATANTE_NUMERO = os.getenv('CONTRATANTE_NUMERO')
CERT_FILENAME      = os.getenv('CERT_PATH')
KEY_FILENAME       = os.getenv('KEY_PATH')
TOKEN_FILE         = os.getenv('TOKEN_FILE')
XML_DIR_NAME       = os.getenv('XML_DIR', 'xml_files')

# 5) Monta **caminhos absolutos** a partir da raiz do projeto
PFX_PATH    = CERTS_DIR / PFX_FILE
CERT_PATH   = CERTS_DIR / CERT_FILENAME
KEY_PATH    = CERTS_DIR / KEY_FILENAME
TOKEN_PATH  = BASE_DIR / TOKEN_FILE
XML_DIR     = BASE_DIR / XML_DIR_NAME

# 6) Debug para verificar se bate certinho
print(f"DEBUG: PFX_PATH   = {PFX_PATH}")
print(f"DEBUG: CERT_PATH  = {CERT_PATH}")
print(f"DEBUG: KEY_PATH   = {KEY_PATH}")
print(f"DEBUG: TOKEN_PATH = {TOKEN_PATH}")
print(f"DEBUG: XML_DIR    = {XML_DIR}")

# Diretorio de trabalho e pasta de XMLs
script_dir    = os.path.dirname(os.path.abspath(__file__))
xml_directory = os.path.join(script_dir, XML_DIR)
os.makedirs(xml_directory, exist_ok=True)

# Usa os Paths absolutos que já definimos lá em cima
cert_path             = CERT_PATH      # Path para certs/certificado.pem
key_path              = KEY_PATH       # Path para certs/chave_privada.pem
caminho_arquivo_token = TOKEN_PATH     # Path para token_info.txt na raiz




def salvar_xml_base64(base64_xml, file_name, output_path):
    """Decodifica um XML em base64 e salva na pasta escolhida pelo usuário."""
    try:
        xml_decodificado = base64.b64decode(base64_xml).decode("utf-8")
        file_path = os.path.join(output_path, file_name)

        with open(file_path, "w", encoding="utf-8") as file:
            file.write(xml_decodificado)

        print(f"✅ XML salvo: {file_path}")
        return file_path
    except Exception as e:
        print(f"❌ Erro ao salvar XML: {e}")
        return None


def processar_xml_para_excel(output_path):
    """Processa os arquivos XML na pasta do usuário e gera planilhas Excel no mesmo local."""
    namespace = {"ns": "http://www.serpro.gov.br/dctf/v1"}

    resumo_contribuintes = []
    tributos_detalhados = []

    # Listando os arquivos XML no diretório escolhido
    files = [f for f in os.listdir(output_path) if f.lower().endswith(".xml")]
    if not files:
        print("⚠️ Nenhum arquivo XML encontrado na pasta escolhida.")
        return

    print(f"📂 {len(files)} arquivos XML encontrados: {files}")

    # Percorre todos os arquivos XML no diretório
    for filename in files:
        file_path = os.path.join(output_path, filename)
        print(f"\n📄 Processando arquivo: {filename}")

        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            print(f"✅ Arquivo {filename} carregado com sucesso!")
        except ET.ParseError as e:
            print(f"❌ ERRO ao processar {filename}: {e}")
            continue

        # Extraindo informações do contribuinte
        nome_contribuinte = root.find(".//ns:nomeContribuinte", namespace)
        insc_contrib = root.find(".//ns:inscContrib", namespace)

        if nome_contribuinte is None or insc_contrib is None:
            print(f"⚠️ Dados do contribuinte não encontrados no arquivo: {filename}")
            continue

        nome_contribuinte = nome_contribuinte.text if nome_contribuinte is not None else "Desconhecido"
        insc_contrib = insc_contrib.text if insc_contrib is not None else "00000000000000"

        print(f"➡️ Contribuinte: {nome_contribuinte}, Inscrição: {insc_contrib}")

        total_saldo_pagar = 0
        tributos = root.findall(".//ns:A050-CreditosTributariosApurados/ns:CreditoTributarioApurado", namespace)

        if not tributos:
            print(f"⚠️ Nenhum tributo encontrado no arquivo: {filename}. Registrando como 'Sem Débitos'.")
            
            # Adiciona ao controle de tributos indicando "Sem Débitos"
            tributos_detalhados.append([
                nome_contribuinte, insc_contrib, "N/A", "Sem Débitos", 
                "N/A", "N/A", "0.00", "N/A", "0.00", "0.00"
            ])
            
            # Adiciona ao resumo de contribuintes com saldo zero
            resumo_contribuintes.append([nome_contribuinte, insc_contrib, 0.00])
            continue

        # Percorre os tributos
        for credito in tributos:
            cod_receita = credito.find("ns:codReceita", namespace).text if credito.find("ns:codReceita", namespace) is not None else ""
            ct_descricao_tributo = credito.find("ns:ctDescricaoTributo", namespace).text if credito.find("ns:ctDescricaoTributo", namespace) is not None else ""
            ct_cod_grupo = credito.find("ns:ctCodGrupo", namespace).text if credito.find("ns:ctCodGrupo", namespace) is not None else ""
            ct_desc_grupo = credito.find("ns:ctDescGrupo", namespace).text if credito.find("ns:ctDescGrupo", namespace) is not None else ""
            ct_valor = credito.find("ns:ctValor", namespace).text if credito.find("ns:ctValor", namespace) is not None else "0"
            pa_debito = credito.find("ns:paDebito", namespace).text if credito.find("ns:paDebito", namespace) is not None else ""
            vl_total_cred = credito.find("ns:vlTotalCred", namespace).text if credito.find("ns:vlTotalCred", namespace) is not None else "0"
            saldo_a_pagar = credito.find("ns:saldoaPagar", namespace).text if credito.find("ns:saldoaPagar", namespace) is not None else "0"

            saldo_a_pagar_float = float(saldo_a_pagar.replace(",", ".")) if saldo_a_pagar else 0.0
            total_saldo_pagar += saldo_a_pagar_float

            tributos_detalhados.append([
                nome_contribuinte, insc_contrib, cod_receita, ct_descricao_tributo, 
                ct_cod_grupo, ct_desc_grupo, ct_valor, pa_debito, vl_total_cred, saldo_a_pagar
            ])

        resumo_contribuintes.append([nome_contribuinte, insc_contrib, total_saldo_pagar])

    # Criar DataFrames
    df_resumo = pd.DataFrame(resumo_contribuintes, columns=["Nome Contribuinte", "Inscrição", "Total Saldo a Pagar"])
    df_tributos = pd.DataFrame(tributos_detalhados, columns=[
        "Nome Contribuinte", "Inscrição", "Código Receita", "Descrição Tributo", 
        "Código Grupo", "Descrição Grupo", "Valor Tributo", "Período Apuração", 
        "Valor Crédito", "Saldo a Pagar"
    ])

    # Salvar os arquivos Excel na pasta do usuário
    df_resumo.to_excel(os.path.join(output_path, "resumo_contribuintes.xlsx"), index=False)
    df_tributos.to_excel(os.path.join(output_path, "controle_tributos.xlsx"), index=False)

    print("\n📂 Arquivos Excel gerados com sucesso!")
    print(f"📁 Local: {output_path}")



from pathlib import Path

def autenticar_com_serpro_e_obter_tokens():
    url = "https://autenticacao.sapi.serpro.gov.br/authenticate"
    # pega o valor cru do .env

    print(f"DEBUG: usando PFX em → {PFX_PATH}")
    if not PFX_PATH.exists():
        raise FileNotFoundError(f"PFX não encontrado em: {PFX_PATH}")

    # credenciais base64
    cred = base64.b64encode(f"{CONSUMER_KEY}:{CONSUMER_SECRET}".encode()).decode()
    headers = {
        "Authorization": f"Basic {cred}",
        "role-type": "TERCEIROS",
        "content-type": "application/x-www-form-urlencoded"
    }
    body = {'grant_type': 'client_credentials'}

    try:
        r = post(
            url,
            data=body,
            headers=headers,
            verify=True,
            pkcs12_filename=str(PFX_PATH),
            pkcs12_password=PFX_PASSWORD
        )
        r.raise_for_status()
        data = r.json()
        print("✅ Autenticação bem-sucedida.")
        return data['access_token'], data.get('jwt_token', ''), data['expires_in']
    except Exception as e:
        print(f"❌ Erro na autenticação Serpro: {e}")
        return None, None, None








def verificar_indZerada_no_xml(xml_base64):
    """
    Verifica se o XML contém a tag <indZerada> com valor '0'.
    """
    try:
        xml_decodificado = base64.b64decode(xml_base64).decode('utf-8')
        root = ET.fromstring(xml_decodificado)

        # O namespace do documento XML
        namespaces = {
            'ns': 'http://www.serpro.gov.br/dctf/v1'
        }

        # Ajuste o caminho para considerar o namespace
        indZerada = root.find('.//ns:indZerada', namespaces)

        if indZerada is not None:
            print(f"Tag indZerada encontrada com valor: {indZerada.text}")
            return indZerada.text == '0'
        else:
            print("Tag indZerada não encontrada.")
            return False
    except Exception as e:
        print(f"Erro ao analisar o XML: {e}")
        return False




def verificar_signature_value_no_xml(xml_base64):
    """
    Verifica se a tag SignatureValue está presente no XML decodificado de Base64.
    
    :param xml_base64: String com o XML codificado em Base64.
    :return: Bool indicando a presença da tag e o valor da tag se presente.
    """
    try:
        # Decodifica o XML de Base64
        xml_decodificado = base64.b64decode(xml_base64).decode('utf-8')

        # Parseia o XML
        root = ET.fromstring(xml_decodificado)

        # Define o namespace para a busca
        namespaces = {'ds': 'http://www.w3.org/2000/09/xmldsig#'}  # Adiciona mais namespaces conforme necessário

        # Procura pela tag SignatureValue usando o namespace
        signature_value = root.find('.//ds:SignatureValue', namespaces)

        if signature_value is not None:
            print("Tag SignatureValue encontrada.")
            
            return True, signature_value.text
        else:
            print("Tag SignatureValue não encontrada.")
            return False, None
    except Exception as e:
        print(f"Erro ao analisar o XML: {e}")
        return False, None


def autenticar_e_salvar_token():
    print(f"DEBUG: salvando token em → {caminho_arquivo_token}")
    try:
        access_token, jwt_token, expires_in = autenticar_com_serpro_e_obter_tokens()
        if access_token and jwt_token:
            expiracao_ts = time.time() + expires_in - 60
            with open(caminho_arquivo_token, 'w') as f:
                f.write(f"{access_token},{jwt_token},{expiracao_ts}")
            print("✅ Token salvo com sucesso.")
        else:
            print("❌ Falha ao obter tokens.")
    except Exception as e:
        print(f"❌ Erro ao tentar autenticar e salvar o token: {e}")

def verificar_e_obter_token():
    try:
        with open(caminho_arquivo_token, 'r') as f:
            at, jwt, exp = f.read().split(',')
        if time.time() < float(exp):
            return at, jwt
        print("⚠️ Token expirado. Renovando…")
    except Exception:
        print("🔑 Token não encontrado ou inválido. Autenticando…")

    autenticar_e_salvar_token()

    try:
        with open(caminho_arquivo_token, 'r') as f:
            at, jwt, exp = f.read().split(',')
        return at, jwt
    except Exception:
        print(f"❌ Falha crítica: não foi possível criar ou ler {caminho_arquivo_token}")
        sys.exit(1)


# Função melhorada para carregar o DataFrame e retornar ele inteiro
def ler_cnpjs_de_excel(caminho_arquivo):
    # Lê o arquivo Excel
    df = pd.read_excel(caminho_arquivo, dtype={'CNPJ': str})
    # Certifica-se de que a coluna 'Status XML' existe no DataFrame
    if 'Status Consulta XML' not in df.columns:
        df['Status Consulta XML'] = None
    
    df['CNPJ'] = df['CNPJ'].apply(lambda x: str(x).zfill(14))
    
    return df


def consultar_api_CONSXMLDECLARACAO38(access_token, jwt_token, cnpj_contribuinte, ano_pa, mes_pa, categoria, output_path):
    url = "https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Consultar"

    # Obtém um token válido antes da requisição
    access_token, jwt_token = verificar_e_obter_token()

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "jwt_token": f"{jwt_token}"
    }

    dados_dict = {
        "categoria": categoria,
        "anoPA": ano_pa
    }
    if mes_pa:
        dados_dict["mesPA"] = mes_pa

    payload = {
        "contratante": {"numero": CONTRATANTE_NUMERO, "tipo": 2},
        "autorPedidoDados": {"numero": CONTRATANTE_NUMERO, "tipo": 2},
        "contribuinte": {"numero": str(cnpj_contribuinte), "tipo": 2},
        "pedidoDados": {
            "idSistema": "DCTFWEB",
            "idServico": "CONSXMLDECLARACAO38",
            "versaoSistema": "1.0",
            "dados": json.dumps(dados_dict)
        }
    }

    response = requests.post(url, headers=headers, data=json.dumps(payload))
    time.sleep(3)
    
    if response.status_code == 200:
        print(f"✅ Resposta bem-sucedida da API ({response.status_code})")

        try:
            response_data = response.json()  # Converte a resposta para JSON
            print(f"📩 Resposta JSON Recebida: {json.dumps(response_data, indent=4, ensure_ascii=False)}")

            if "dados" in response_data:
                if isinstance(response_data["dados"], str):
                    try:
                        dados = json.loads(response_data["dados"])  # Desserializa 'dados'
                        xml_base64 = dados.get("XMLStringBase64", "")

                        if xml_base64:
                            print("✅ XML encontrado no retorno da API!")

                            # Salva o XML localmente
                            xml_file = salvar_xml_base64(xml_base64, f"{cnpj_contribuinte}_DCTFWEB.xml", output_path)

                            # Se o XML foi salvo com sucesso, processa para Excel
                            if xml_file:
                                processar_xml_para_excel(output_path)

                            return xml_base64
                        else:
                            print("⚠️ 'XMLStringBase64' não encontrado no JSON retornado.")
                    except json.JSONDecodeError as e:
                        print(f"❌ Erro ao decodificar 'dados': {e}")
                else:
                    print("⚠️ O campo 'dados' não está no formato esperado (string JSON).")
            else:
                print("⚠️ O JSON não contém a chave esperada: 'dados'.")
        except json.JSONDecodeError as e:
            print(f"❌ Erro ao interpretar a resposta da API como JSON: {e}")

        return None
    else:
        print(f"❌ Erro na resposta da API: {response.status_code}")
        print(f"📩 Resposta completa: {response.text}")

        # Exibir informações úteis para debugging
        print(f"🔍 Headers Enviados: {json.dumps(headers, indent=4, ensure_ascii=False)}")
        print(f"📦 Payload Enviado: {json.dumps(payload, indent=4, ensure_ascii=False)}")

        return None




def declarar_api_TRANSDECLARACAO310(access_token, jwt_token, cnpj_contribuinte, ano_pa, mes_pa, categoria, xml_assinado_base64):
    url = "https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Declarar"
    # Obtém um token válido antes da requisição
    access_token, jwt_token = verificar_e_obter_token()

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "jwt_token": f"{jwt_token}"
    }

    dados_dict = {
        "categoria": categoria,
        "anoPA": ano_pa,
        "xmlAssinadoBase64": xml_assinado_base64
    }
    if mes_pa:
        dados_dict["mesPA"] = mes_pa

    payload = {
        "contratante": {"numero": "56018088000123", "tipo": 2},
        "autorPedidoDados": {"numero": "56018088000123", "tipo": 2},
        "contribuinte": {"numero": cnpj_contribuinte, "tipo": 2},
        "pedidoDados": {
            "idSistema": "DCTFWEB",
            "idServico": "TRANSDECLARACAO310",
            "versaoSistema": "1.0",
            "dados": json.dumps(dados_dict)
        }
    }
    # Imprime o payload antes de enviar
    print("Payload sendo enviado:")
    print(json.dumps(payload, indent=4))  # Usa indentação para melhor visualização

    response = requests.post(url, headers=headers, data=json.dumps(payload))
    time.sleep(3)
    
    if response.status_code == 200:
        print(f"Resposta bem-sucedida da API: {response.status_code}")
        print(f"Resposta: {response.text}")
    else:
        print(f"Erro na resposta da API: {response.status_code}")
        print(f"Resposta: {response.text}")

    return response



def assinar_xml(xml_base64, cert_path, key_path):
    try:
        xml_bytes = base64.b64decode(xml_base64)
        tree = etree.fromstring(xml_bytes)

        # Encontra o elemento ConteudoDeclaracao e obtém o valor do atributo id
        conteudo_decl = tree.find(".//{http://www.serpro.gov.br/dctf/v1}ConteudoDeclaracao")
        if conteudo_decl is not None:
            ref_uri = f"#{conteudo_decl.get('id')}"
        else:
            raise ValueError("Elemento ConteudoDeclaracao não encontrado")

        with open(cert_path, 'rb') as f_cert, open(key_path, 'rb') as f_key:
            cert = f_cert.read()
            key = f_key.read()

        signer = XMLSigner(method=methods.enveloped, signature_algorithm='rsa-sha256',
                           digest_algorithm='sha256', c14n_algorithm='http://www.w3.org/TR/2001/REC-xml-c14n-20010315')

        signed_root = signer.sign(tree, key=key, cert=cert, reference_uri=ref_uri)

        signed_xml = etree.tostring(signed_root, encoding='utf-8', xml_declaration=True)
        signed_xml_base64 = base64.b64encode(signed_xml).decode('utf-8')

        return signed_xml_base64
    except Exception as e:
        print(f"Erro ao assinar o XML: {e}")
        return None


def emitir_api_GERARGUIA31(access_token, jwt_token, cnpj_contribuinte, output_path, ano_pa, mes_pa, categoria):
    url = "https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Emitir"

        # Obtém um token válido antes da requisição
    access_token, jwt_token = verificar_e_obter_token()

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "jwt_token": jwt_token
    }

    dados_dict = {
        "categoria": categoria,
        "anoPA": ano_pa
    }
    if mes_pa:
        dados_dict["mesPA"] = mes_pa

    payload = {
        "contratante": {"numero": "56018088000123", "tipo": 2},
        "autorPedidoDados": {"numero": "56018088000123", "tipo": 2},
        "contribuinte": {"numero": cnpj_contribuinte, "tipo": 2},
        "pedidoDados": {
            "idSistema": "DCTFWEB",
            "idServico": "GERARGUIA31",
            "versaoSistema": "1.0",
            "dados": json.dumps(dados_dict)
        }
    }

    response = requests.post(url, headers=headers, json=payload)
    time.sleep(3)

    if response.status_code == 200:
        resposta_json = response.json()
        mensagens = resposta_json.get("mensagens", [])
        # Verifica se há mensagens de sucesso conforme esperado
        if any(mensagem["codigo"] == "[Sucesso-DCTFWEB]" for mensagem in mensagens):
            if resposta_json.get("dados"):
                pdf_base64 = json.loads(resposta_json["dados"]).get("PDFByteArrayBase64", "")
                if pdf_base64:
                    pdf_bytes = base64.b64decode(pdf_base64)
                    nome_arquivo_pdf = f"{cnpj_contribuinte}_guia_pagamento.pdf"
                    caminho_completo_pdf = os.path.join(output_path, nome_arquivo_pdf)
                    with open(caminho_completo_pdf, "wb") as pdf_file:
                        pdf_file.write(pdf_bytes)
                    print(f"Guia de pagamento emitida e salva com sucesso em: {caminho_completo_pdf}.")
                    return "Sucesso"
                else:
                    # Se não houver PDF, retorna a resposta completa da API para análise
                    print("Erro: A resposta da API não contém dados de PDF. Resposta completa:", resposta_json)
                    return f"Erro: Sem PDF. Resposta: {resposta_json}"
            else:
                print("Erro: A resposta da API não contém dados de PDF. Resposta completa:", resposta_json)
                return f"Erro: Sem dados. Resposta: {resposta_json}"
        else:
            print("Erro: A resposta da API indica que a emissão não foi bem-sucedida. Resposta completa:", resposta_json)
            return f"Erro: Emissão falhou. Resposta: {resposta_json}"
    else:
        error_message = f"Falha ao emitir a guia de pagamento. Código de status: {response.status_code}, Resposta da API: {response.text}"
        print(error_message)
        return f"Erro: {error_message}"
    
    

def processar_arquivo(excel_path, output_path, script_dir, ano_pa, mes_pa, categoria):
    df = ler_cnpjs_de_excel(excel_path)
    xmls_base64_por_cnpj = {}

    for index, row in df.iterrows():
        cnpj = row['CNPJ']
        print(f"🔍 Consultando API para o CNPJ: {cnpj}")

        # Obtém um token válido ANTES de cada requisição
        access_token, jwt_token = verificar_e_obter_token()

        xml_base64 = xmls_base64_por_cnpj.get(cnpj)
        
        if not xml_base64:
            xml_base64 = consultar_api_CONSXMLDECLARACAO38(access_token, jwt_token, cnpj, ano_pa, 
                                                            mes_pa if categoria == "GERAL_MENSAL" else None, 
                                                            categoria, output_path)
            xmls_base64_por_cnpj[cnpj] = xml_base64

        tag_encontrada, valor_signature = verificar_signature_value_no_xml(xml_base64)
        
        if tag_encontrada:
            print(f"✅ XML já assinado para o CNPJ {cnpj}. Emitindo guia de pagamento.")

            # Obtém um token válido antes de emitir a guia
            access_token, jwt_token = verificar_e_obter_token()

            status_guia = emitir_api_GERARGUIA31(access_token, jwt_token, cnpj, output_path, ano_pa, mes_pa, categoria)
            df.at[index, 'Status Consulta XML'] = "DCTFWEB transmitida anteriormente"
            df.at[index, 'Status GUIA'] = status_guia
        else:
            print("✍️ Assinando o XML, pois a tag SignatureValue não foi encontrada.")
            xml_assinado_base64 = assinar_xml(xml_base64, cert_path, key_path)

            if xml_assinado_base64:
                # Obtém um token válido antes de declarar
                access_token, jwt_token = verificar_e_obter_token()

                response = declarar_api_TRANSDECLARACAO310(access_token, jwt_token, cnpj, ano_pa, mes_pa, categoria, xml_assinado_base64)
                if response and response.status_code == 200:
                    print(f"✅ Declaração transmitida com sucesso para o CNPJ {cnpj}.")
                    df.at[index, 'Status Consulta XML'] = "Transmitido DCTFWEB via API"

                    # Obtém um token válido antes de emitir a guia
                    access_token, jwt_token = verificar_e_obter_token()

                    print(f"📄 Emitindo guia de pagamento para o CNPJ {cnpj}")
                    status_guia = emitir_api_GERARGUIA31(access_token, jwt_token, cnpj, output_path, ano_pa, mes_pa, categoria)
                    df.at[index, 'Status GUIA'] = status_guia
                else:
                    print(f"❌ Erro na transmissão para o CNPJ {cnpj}.")
                    df.at[index, 'Status Consulta XML'] = "Erro na Transmissão"
                    df.at[index, 'Status GUIA'] = "A resposta da API não contém dados de PDF"
            else:
                print("⚠️ Erro ao assinar o XML.")
                df.at[index, 'Status Consulta XML'] = "Erro na Assinatura"

        time.sleep(4)  # Pausa entre consultas para evitar sobrecarga

    # Atualiza e salva a planilha ao final do processo
    caminho_arquivo_atualizado = os.path.join(output_path, 'LayoutExemplo_Atualizado.xlsx')
    df.to_excel(caminho_arquivo_atualizado, index=False)
    print(f"📊 Arquivo Excel atualizado salvo em: {caminho_arquivo_atualizado}")


def selecionar_arquivo_excel(entry):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    entry.delete(0, tk.END)
    entry.insert(0, filename)

def selecionar_pasta_salvamento(entry):
    foldername = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, foldername)


class StdoutRedirector(object):
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)

    def flush(self):
        pass




def executar(script_dir, excel_path, output_path, ano_pa, mes_pa, categoria, loading_label, output_text):
    if not excel_path or not output_path:
        messagebox.showwarning("Aviso", "Por favor, selecione o arquivo Excel e a pasta de salvamento.")
        return

    loading_label.config(text="Processando...")

    # Cria e inicia uma thread para o processamento
    process_thread = threading.Thread(target=processar_em_thread, args=(script_dir, excel_path, output_path, ano_pa, mes_pa, categoria, loading_label, output_text))
    process_thread.start()

    # Redireciona stdout para a caixa de texto na GUI
    sys.stdout = StdoutRedirector(output_text)


def processar_em_thread(script_dir, excel_path, output_path, ano_pa, mes_pa, categoria, loading_label, output_text):
    try:
        sys.stdout = StdoutRedirector(output_text)  # Redireciona o stdout
        processar_arquivo(excel_path, output_path, script_dir, ano_pa, mes_pa, categoria)
        loading_label.config(text="")  # Limpa o texto de carregamento
        messagebox.showinfo("Sucesso", "Processamento concluído com sucesso.")  # Exibe mensagem de sucesso
    except Exception as e:
        loading_label.config(text="")  # Limpa o texto de carregamento em caso de erro
        messagebox.showerror("Erro", f"Erro durante o processamento: {str(e)}")  # Exibe mensagem de erro


def iniciar_gui(script_dir):
    root = tk.Tk()
    root.title("WeBot - DCTFWEB-V8")

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10)

    # Seleção de Categoria
    tk.Label(frame, text="Tipo de Declaração:").grid(row=0, column=0, sticky="w")
    tipo_var = tk.StringVar(value="GERAL_MENSAL")

    def ajustar_campos():
        if tipo_var.get() == "GERAL_MENSAL":
            # Mostra o campo Mês PA
            label_mes_pa.grid(row=4, column=0, sticky="w")
            entry_mes_pa.grid(row=4, column=1)
        else:
            # Oculta o campo Mês PA
            label_mes_pa.grid_remove()
            entry_mes_pa.grid_remove()

    tk.Radiobutton(frame, text="Mensal", variable=tipo_var, value="GERAL_MENSAL", command=ajustar_campos).grid(row=0, column=1, sticky="w")
    tk.Radiobutton(frame, text="13º Salário", variable=tipo_var, value="GERAL_13o_SALARIO", command=ajustar_campos).grid(row=0, column=2, sticky="w")

    # Campos existentes
    tk.Label(frame, text="Arquivo Excel:").grid(row=1, column=0, sticky="w")
    entry_excel_path = tk.Entry(frame, width=50)
    entry_excel_path.grid(row=1, column=1)
    tk.Button(frame, text="Selecionar", command=lambda: selecionar_arquivo_excel(entry_excel_path)).grid(row=1, column=2)

    tk.Label(frame, text="Pasta de Salvamento:").grid(row=2, column=0, sticky="w")
    entry_output_path = tk.Entry(frame, width=50)
    entry_output_path.grid(row=2, column=1)
    tk.Button(frame, text="Selecionar", command=lambda: selecionar_pasta_salvamento(entry_output_path)).grid(row=2, column=2)

    # Ano PA
    tk.Label(frame, text="Ano PA:").grid(row=3, column=0, sticky="w")
    entry_ano_pa = tk.Entry(frame, width=50)
    entry_ano_pa.grid(row=3, column=1)

    # Mês PA
    label_mes_pa = tk.Label(frame, text="Mês PA:")
    label_mes_pa.grid(row=4, column=0, sticky="w")
    entry_mes_pa = tk.Entry(frame, width=50)
    entry_mes_pa.grid(row=4, column=1)

    # Label para mensagem de carregamento
    loading_label = tk.Label(root, text="")
    loading_label.pack()

    # Caixa de texto para saída
    output_text = scrolledtext.ScrolledText(root, height=10)
    output_text.pack(pady=10)

    tk.Button(
        root,
        text="Executar",
        command=lambda: executar(
            script_dir,
            entry_excel_path.get(),
            entry_output_path.get(),
            entry_ano_pa.get(),
            entry_mes_pa.get() if tipo_var.get() == "GERAL_MENSAL" else None,
            tipo_var.get(),
            loading_label,
            output_text
        )
    ).pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    iniciar_gui(script_dir)