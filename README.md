# WeBot - DCTFWEB-V8

## 1. Sobre o Projeto
O **Webot DCTFWEB - V8** é uma ferramenta em Python com interface gráfica (Tkinter) para automação das operações de:
- Autenticação na API de Integra Contador do SERPRO (DCTFWEB).
- Consulta e download de declarações em formato XML.
- Assinatura digital dos XMLs utilizando certificado PFX/PEM.
- Transmissão das declarações (CONSXMLDECLARACAO38 e TRANSDECLARACAO310).
- Geração de guias de pagamento (GERARGUIA31) e extração de PDF.
- Processamento de arquivos XML para planilhas Excel com resumo de contribuintes e controle de tributos.

## 2. Integra Contador (Credenciais)
Para se comunicar com a API “Integra Contador” do Serpro, crie um arquivo `.env` na raiz do projeto com estas variáveis:

    PFX_FILE=seu_certificado.pfx
    PFX_PASSWORD=sua_senha_do_pfx
    CONSUMER_KEY=sua_consumer_key
    CONSUMER_SECRET=seu_consumer_secret
    CONTRATANTE_NUMERO=seu_cnpj_sem_mascara   # Ex: 12345678000123
    CERT_PATH=certs/certificado.pem
    KEY_PATH=certs/chave_privada.pem
    TOKEN_FILE=token_info.txt
    XML_DIR=xml_files

- **PFX_FILE** / **PFX_PASSWORD**: seu certificado digital em PKCS#12.  
- **CONSUMER_KEY** / **CONSUMER_SECRET**: credenciais OAuth do Serpro.  
- **CONTRATANTE_NUMERO**: CNPJ do contribuinte (somente números).  
- **CERT_PATH** / **KEY_PATH**: caminhos para o certificado e chave em PEM.  
- **TOKEN_FILE**: onde o token OAuth será salvo.  
- **XML_DIR**: pasta de saída dos XMLs.

## 3. Certificado Digital e Conversão para PEM
O WeBot requer **certificado** e **chave privada** em formato PEM. Se você instalou o **OpenSSL** (ou outro utilitário PKI) na sua máquina, faça:

1. **Exportar somente o certificado** (sem chave privada):

        openssl pkcs12 \
          -in certs/seu_certificado.pfx \
          -clcerts -nokeys \
          -out certs/certificado.pem \
          -passin pass:SUA_SENHA_DO_PFX

2. **Exportar somente a chave privada** (sem certificado):

        openssl pkcs12 \
          -in certs/seu_certificado.pfx \
          -nocerts -nodes \
          -out certs/chave_privada.pem \
          -passin pass:SUA_SENHA_DO_PFX

### Estrutura final do diretório `certs/`
    certs/
    ├── seucertificado.pfx
    ├── certificado.pem
    ├── chave_privada.pem
    └── .gitignore          # ou .gitkeep para manter a pasta vazia no Git
