import win32com.client as win32
import pandas as pd
# Integração
outlook = win32.Dispatch('outlook.application')
caminho_tabela = "emails.xlsx"

aba_planilha = "planilha 1"
tabela_lida = pd.read_excel(caminho_tabela, sheet_name=aba)

objeto_busca = {
    "planilha 1": "http://link_a_ser_enviado",
}

link_a_enviar = objeto_busca[aba]

# Configurar email

# email.To = "destino"
# email.Subject = "assunto"
# email.HTMLBody = "Corpo do E-mail"
# list_to_emails = tabela.values.tolist()
# print(list_to_emails)
contador = 0
for indice, linha in tabela_lida.iterrows():
    email = ""
    email = outlook.CreateItem(0)
    print(linha.iloc[0])
    nome = linha.iloc[0]
    print(linha.iloc[1])
    email.To = linha.iloc[1]
    email.Subject = "Link enviado"

    email.HTMLBody = f''''
        <!DOCTYPE html>
        <html>
            <head>
                <style>
                    *{{
                        margin: 0px;
                        padding: 0px;
                    }}
                    
                    .main{{
                        font-family: sans-serif;
                        text-align: center;
                        color: white;
                        border-radius: 20px;
                        border: 10px solid #F8AA36;
                        margin: 10px;
                    }}
                    .cabecalho{{
                        background-color: #855CD6;
                        font-size: 30px;
                        padding: 50px;
                    }}
                    .conteudo{{
                        height: 400px;
                        background-color: #855CD6;
                    }}
                    #titulo{{
                        color: #F8AA36;

                    }}
                    p {{
                        font-size: 40px;
                        font-family: sans-serif;
                        color: floralwhite;
                        padding: 40px;

                        
                    }}
                    #pronto{{
                        color: white;
                        background-color: #855CD6;
                        text-decoration: none;
                        font-size: 40px;
                        padding: 10px 20px 10px 20px;
                        border-radius: 10px;
                    }}
                    #pronto:hover{{
                        background-color: #7753C0;
                    }}
                </style>
                
            </head>
            <body>
                <div class="main">
                    <header class="cabecalho">
                        <h1 id="titulo">Bem-vindo(a) {nome}</h1> 
                    </header>
                    <div class="conteudo">
                        <p>Pronto(a) para essa jornada no mundo da programação?</p>
                        <br>
                        <br>
                        <b><a id="pronto" href="{link_a_enviar}">Acessar Scratch</a></b>
                    </div>

                </div>

            </body>
        </html>
    '''




    contador += 1
    email.Send()
    print(f"Email nº{contador} enviado")
