📩 Automação de Relatórios e Envio de E-mails com Python
Este projeto automatiza a consulta de dados de produção a partir de um banco de dados SQL Server, gera um relatório em Excel e o envia por e-mail utilizando Python.

🚀 Tecnologias Utilizadas
Python 3.12.1
SQLAlchemy (Conexão com o SQL Server)
Pandas (Manipulação de dados)
smtplib (Envio de e-mails)
MIME (Anexação de arquivos ao e-mail)
📌 Funcionalidades
✅ Conexão segura com SQL Server usando SQLAlchemy
✅ Extração automática da última data disponível na base de produção
✅ Filtros para selecionar os registros desejados
✅ Geração automática de um relatório em Excel
✅ Envio do relatório por e-mail de forma automatizada

🔧 Como Usar
1️⃣ Configurar Banco de Dados
O código usa SQLAlchemy para conectar ao banco de dados. Altere a string de conexão conforme necessário:

python
Copiar
Editar
engine = create_engine('mssql+pyodbc://USUARIO:SENHA@SERVIDOR/NOME_DO_BD?driver=SQL+Server')
2️⃣ Configurar Envio de E-mails
No código, altere os detalhes do e-mail para o seu próprio remetente:

python
Copiar
Editar
from_email = 'seuemail@gmail.com'  
senha_email = 'SUA_SENHA_DE_APP'  
to_emails = ['destinatario@email.com']  
Importante: Para enviar e-mails pelo Gmail, você precisa gerar uma Senha de App. Veja como fazer isso aqui.

3️⃣ Executar o Script
Basta rodar o script no terminal:

bash
Copiar
Editar
python automacao_relatorio.py
📂 Estrutura do Projeto
bash
Copiar
Editar
📂 projeto-relatorio
│── 📄 automacao_relatorio.py  # Script principal
│── 📄 README.md               # Documentação do projeto
│── 📄 RELATORIO.xlsx          # Relatório gerado (será substituído a cada execução)
🔥 Próximos Passos
🔹 Melhorar a formatação do relatório Excel
🔹 Criar uma interface web com Streamlit
🔹 Integrar a automação com uma API

👨‍💻 Autor
Projeto desenvolvido por Eduardo Oliveira. 🚀

Se este projeto foi útil para você, não se esqueça de dar uma ⭐ no GitHub! 😃
