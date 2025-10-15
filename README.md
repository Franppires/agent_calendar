# agent_calendar
# Agente de Calendário Proativo

Este aplicativo agenda eventos no Google Calendar de forma proativa, com interface gráfica.

## Como usar

1. **Clone o repositório:**
	```
	git clone https://github.com/Franppires/agent_calendar.git
	cd agent_calendar
	```

2. **Instale o Python (recomendado 3.10+).**

3. **Instale as dependências:**
	```
	pip install -r requirements.txt
	```

4. **Crie o arquivo de credenciais:**
	- Acesse [Google Cloud Console](https://console.cloud.google.com/).
	- Crie um projeto, ative a API do Google Calendar.
	- Crie credenciais do tipo "OAuth Client ID" para aplicativo de desktop.
	- Baixe o arquivo `credentials.json` e coloque na pasta principal do projeto.

5. **Execute o aplicativo:**
	```
	python calendar_agent.py
	```

## Observações

- O arquivo `credentials.json` **não** deve ser compartilhado publicamente.
- O app irá pedir autenticação Google na primeira execução.
- Os eventos serão criados no calendário selecionado.
