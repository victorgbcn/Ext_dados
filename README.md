# Ext_dados
Automacao de Extracao de Dados - Magazine Luiza

Descricao

Este projeto automatiza a extracao de informacoes sobre notebooks no site da Magazine Luiza. A automacao acessa o site, pesquisa por "notebooks", extrai os dados das cinco primeiras paginas e gera um relatorio em Excel. O relatorio e enviado por e-mail ao usuario.

Tecnologias Utilizadas

Python

Selenium

Pandas

PyAutoGUI

SMTP (para envio de e-mail)

Requisitos

Antes de rodar o script, instale as dependencias executando:

pip install -r requirements.txt

O arquivo requirements.txt inclui as seguintes dependências:

selenium
pandas
pyautogui
openpyxl
python-dotenv

Certifique-se de ter o Google Chrome instalado e o ChromeDriver compatível com sua versão do navegador.
Antes de rodar o script, instale as dependencias executando:

pip install -r requirements.txt

Certifique-se de ter o Google Chrome instalado e o ChromeDriver compatível com sua versão do navegador.

Configuracao

Variaveis de Ambiente:

Crie um arquivo .env e adicione sua senha de e-mail:

EMAIL_SENHA=minha_senha_aqui

No script, use os.getenv("EMAIL_SENHA") para acessar a senha de forma segura.

Execucao:

Rode o script com:

python automacao.py

Estrutura do Projeto

├── automacao.py         # Codigo principal
├── requirements.txt     # Dependencias do projeto
├── .env                 # Variaveis de ambiente (NÃO COMMITAR!)
├── Teste Pratico planilha/ # Pasta onde a planilha sera salva

Melhorias Futuras

Extracao de mais dados (preco, descricao, etc.)

Implementacao de um banco de dados para armazenamento dos resultados

Melhor tratamento de erros e logs

Autor

Victor Gabriel Caldeira Nicolau


