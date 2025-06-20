import gdown
import pandas as pd
import yagmail

# === ETAPA 1: Baixar o arquivo do Google Drive ===
# Link compartilhado
url = 'https://drive.google.com/uc?id=1GBeZPW9iXnqAmNyN40GSouUROuBNr-3i'
output = 'tendencia.xlsx'

# Baixando o arquivo
gdown.download(url, output, quiet=False)

# === ETAPA 2: Ler planilhas ===
xls = pd.ExcelFile(output)

# Planilha "DIAS_TRABALHO"
df_dias = pd.read_excel(xls, sheet_name='DIAS_TRABALHO').iloc[2:5]
df_dias.columns = ['_', 'Indicador', 'Valor']
df_dias = df_dias[['Indicador', 'Valor']].dropna()

# Planilha "CONTROLE"
df_ctrl = pd.read_excel(xls, sheet_name='CONTROLE').iloc[3:].reset_index(drop=True)
df_ctrl.columns = pd.read_excel(xls, sheet_name='CONTROLE').iloc[2]
df_ctrl = df_ctrl[df_ctrl['PILARES'].notna()]
df_ctrl = df_ctrl[df_ctrl['PILARES'] != 'TOTAL']

# Conversões
df_ctrl[['REAL', 'META', 'Projeção']] = df_ctrl[['REAL', 'META', 'Projeção']].apply(pd.to_numeric, errors='coerce')
df_ctrl[['% Real Meta', '% Proj da Meta']] = df_ctrl[['% Real Meta', '% Proj da Meta']].apply(pd.to_numeric, errors='coerce')

# === ETAPA 3: Gerar análise ===
melhor = df_ctrl.loc[df_ctrl['% Real Meta'].idxmax()]
pior = df_ctrl.loc[df_ctrl['% Real Meta'].idxmin()]
proj = df_ctrl.loc[df_ctrl['% Proj da Meta'].idxmax()]

print(f"""
🔍 ANÁLISE AUTOMÁTICA

➡️ MELHOR PILAR: {melhor['PILARES']}
   Realizado: R$ {melhor['REAL']:,.2f} | Meta: R$ {melhor['META']:,.2f}
   Percentual da Meta: {melhor['% Real Meta']:.2%}

➡️ PIOR PILAR: {pior['PILARES']}
   Realizado: R$ {pior['REAL']:,.2f} | Meta: R$ {pior['META']:,.2f}
   Percentual da Meta: {pior['% Real Meta']:.2%}

📈 MELHOR PROJEÇÃO: {proj['PILARES']}
   Projeção: R$ {proj['Projeção']:,.2f}
   Percentual projetado: {proj['% Proj da Meta']:.2%}
""")

# Texto gerado no final da análise
mensagem = f"""
🔍 ANÁLISE AUTOMÁTICA

➡️ MELHOR PILAR: {melhor['PILARES']}
   Realizado: R$ {melhor['REAL']:,.2f} | Meta: R$ {melhor['META']:,.2f}
   Percentual da Meta: {melhor['% Real Meta']:.2%}

➡️ PIOR PILAR: {pior['PILARES']}
   Realizado: R$ {pior['REAL']:,.2f} | Meta: R$ {pior['META']:,.2f}
   Percentual da Meta: {pior['% Real Meta']:.2%}

📈 MELHOR PROJEÇÃO: {proj['PILARES']}
   Projeção: R$ {proj['Projeção']:,.2f}
   Percentual projetado: {proj['% Proj da Meta']:.2%}
"""

# === Configuração do email ===
remetente = 'apollolds2@gmail.com'
senha = 'bmxf pgaq ezyp yvon'
destinatario = 'apollolopeeees@gmail.com'

# Enviar e-mail
yag = yagmail.SMTP(remetente, senha)
yag.send(
    to=destinatario,
    subject='Análise de Tendências - Relatório Automático',
    contents=mensagem
)

print("📧 E-mail enviado com sucesso!")