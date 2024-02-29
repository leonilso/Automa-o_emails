import win32com.client as win32
import pandas as pd
# Integração
outlook = win32.Dispatch('outlook.application')
caminho_tabela = "emails.xlsx"

# turma = "teste"
# turma = "primeiro_ano"
# turma = "oitavo_ano_A_prata"
# turma = "oitavo_ano_B_prata"
turma = "nono_ano_A_prata"
# turma = "nono_ano_B_prata"
# turma = "oitavo_ano_guarani"
# turma = "nono_ano_guarani"
tabela = pd.read_excel(caminho_tabela, sheet_name=turma)

primeiro_ano = "http://scratch.mit.edu/signup/n25yd5n3n"
oitavo_ano_A_prata = "http://scratch.mit.edu/signup/m4kenv8rf"
oitavo_ano_B_prata = "http://scratch.mit.edu/signup/vh8h3pchc"
nono_ano_A_prata = "http://scratch.mit.edu/signup/yy4w9hftc"
nono_ano_B_prata = "http://scratch.mit.edu/signup/39yytt489"
oitavo_ano_guarani = "http://scratch.mit.edu/signup/33m3fwexv"
nono_ano_guarani = "http://scratch.mit.edu/signup/pfc5mdr5d"

if turma == "oitavo_ano_A_prata":
    link_scratch = nono_ano_A_prata
elif turma == "oitavo_ano_B_prata":
    link_scratch = oitavo_ano_B_prata
elif turma == "oitavo_ano_ano_guarani":
    link_scratch = oitavo_ano_guarani
elif turma == "nono_ano_A_prata":
    link_scratch = nono_ano_A_prata
elif turma == "nono_ano_B_prata":
    link_scratch = nono_ano_B_prata
elif turma == "nono_ano_guarani":
    link_scratch = nono_ano_guarani
elif turma == "primeiro_ano":
    link_scratch = primeiro_ano
else:
    link_scratch = primeiro_ano





# nome_arquivo = "foto.jpg"
# caminho_foto = os.path.abspath(nome_arquivo)

# # Anexando a foto
# attachment = email.Attachments.Add(caminho_foto)
# attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "foto")

# Configurar email

# email.To = "destino"
# email.Subject = "assunto"
# email.HTMLBody = "Corpo do E-mail"
# list_to_emails = tabela.values.tolist()
# print(list_to_emails)
contador = 0
for indice, linha in tabela.iterrows():
    email = ""
    email = outlook.CreateItem(0)
    print(linha.iloc[0])
    nome = linha.iloc[0]
    print(linha.iloc[1])
    email.To = linha.iloc[1]
    email.Subject = "Link do Scratch"

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
                        <b><a id="pronto" href="{link_scratch}">Acessar Scratch</a></b>
                    </div>

                </div>

            </body>
        </html>
    '''




    contador += 1
    email.Send()
    print(f"Email nº{contador} enviado")