# WeBot - DCTFWEB

## 1. Sobre o Projeto
O **Webot - DCTFWEB** é uma ferramenta em Python com interface gráfica (Tkinter) para automação das operações de:
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
    ├── seucertificado.pem
    ├── certificado.pem
    ├── chave_privada.pem
    └── .gitignore          # ou .gitkeep para manter a pasta vazia no Git

## 4. Documentação Adicional
Se quiser saber mais detalhes sobre a API Integra Contador, consulte a documentação oficial em:  
https://apicenter.estaleiro.serpro.gov.br/documentacao/api-integra-contador/


## 5. Executando o WeBot
Após configurar o `.env`, siga estes passos:

1. **Preenchimento do layout de exemplo**  
   - Abra o arquivo `LayoutExemplo.xlsx` (na raiz do projeto).  
   - Preencha a coluna **CNPJ** com o CNPJ da empresa que deseja processar (14 dígitos, sem formatação).  
   - Para múltiplas empresas, coloque cada CNPJ em uma linha separada.  
   - Salve o arquivo.

2. **Preencha os demais campos na GUI**  
   - **Arquivo Excel**: selecione o `LayoutExemplo.xlsx` preenchido.  
   - **Pasta de Salvamento**: escolha onde os XMLs e o Excel atualizado serão gerados.  
   - **Ano PA**: informe o ano de competência (ex: `2025`).  
   - **Mês PA**: informe o mês de competência (`01` a `12`; só disponível se “Mensal” estiver selecionado).  
   - **Tipo de Declaração**: escolha entre **Mensal** ou **13º Salário**.

3. **Executar**  
   - Clique em **Executar**.  
   - Acompanhe o log na janela para verificar o andamento.  
   - Ao final, será gerado um arquivo `LayoutExemplo_Atualizado.xlsx` na pasta de saída, contendo os status de cada CNPJ e, se aplicável, as guias de pagamento.


## 6. Empacotando com PyInstaller
Para facilitar o uso por outros colaboradores, gere um executável usando o [PyInstaller](https://pyinstaller.org/):

1. Instale o PyInstaller:
    
        pip install pyinstaller

2. Na raiz do projeto, execute (exemplo Windows):

        pyinstaller --onefile --windowed \
          --add-data "certs;certs" \
          --add-data ".env;." \
          Nucleo-DCTFWEB-Entrega&GeraGuiaV8.py

   - `--add-data "certs;certs"`: inclui toda a pasta `certs/` (PFX e PEMs)  
   - `--add-data ".env;."`: inclui o arquivo de variáveis de ambiente  

3. Após a execução, a pasta `dist/` conterá:
   - O executável do WeBot  
   - A pasta `certs/` com seus certificados  
   - O arquivo `.env`  

4. Compartilhe a pasta `dist/` (ou apenas o executável junto de `certs/` e `.env`) para que outros colaboradores possam rodar sem instalar dependências.