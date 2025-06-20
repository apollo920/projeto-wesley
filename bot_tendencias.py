import gdown
import pandas as pd
import yagmail

def executar_bot():
    # TODO: copie todo seu código aqui
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

    print("🔍 RELATÓRIO DE ANÁLISE AUTOMÁTICA\n")

    # MELHOR PILAR
    print("✅ MELHOR PILAR ATUAL:")
    print(f"   🔹 Pilar: {melhor['PILARES']}")
    print(f"   💰 Realizado: R$ {melhor['REAL']:,.2f}")
    print(f"   🎯 Meta: R$ {melhor['META']:,.2f}")
    print(f"   📊 Percentual da Meta: {melhor['% Real Meta']:.2%}")
    print()

    # PIOR PILAR
    print("❌ PIOR PILAR ATUAL:")
    print(f"   🔸 Pilar: {pior['PILARES']}")
    real_pior = "Indisponível" if pd.isna(pior['REAL']) else f"R$ {pior['REAL']:,.2f}"
    print(f"   💰 Realizado: {real_pior}")
    print(f"   🎯 Meta: R$ {pior['META']:,.2f}")
    print(f"   📉 Percentual da Meta: {pior['% Real Meta']:.2%}")
    print()

    # MELHOR PROJEÇÃO
    print("📈 MELHOR PROJEÇÃO FUTURA:")
    print(f"   🔹 Pilar: {proj['PILARES']}")
    print(f"   📈 Projeção: R$ {proj['Projeção']:,.2f}")
    print(f"   🎯 Percentual projetado: {proj['% Proj da Meta']:.2%}")
    print()

    print("📬 Relatório finalizado com sucesso.")


    # Texto gerado no final da análise
    mensagem = "🔍 RELATÓRIO DE ANÁLISE AUTOMÁTICA\n\n"

    mensagem += "✅ MELHOR PILAR ATUAL:\n"
    mensagem += f"   🔹 Pilar: {melhor['PILARES']}\n"
    mensagem += f"   💰 Realizado: R$ {melhor['REAL']:,.2f}\n"
    mensagem += f"   🎯 Meta: R$ {melhor['META']:,.2f}\n"
    mensagem += f"   📊 Percentual da Meta: {melhor['% Real Meta']:.2%}\n\n"

    mensagem += "❌ PIOR PILAR ATUAL:\n"
    mensagem += f"   🔸 Pilar: {pior['PILARES']}\n"
    real_pior = "Indisponível" if pd.isna(pior['REAL']) else f"R$ {pior['REAL']:,.2f}"
    mensagem += f"   💰 Realizado: {real_pior}\n"
    mensagem += f"   🎯 Meta: R$ {pior['META']:,.2f}\n"
    mensagem += f"   📉 Percentual da Meta: {pior['% Real Meta']:.2%}\n\n"

    mensagem += "📈 MELHOR PROJEÇÃO FUTURA:\n"
    mensagem += f"   🔹 Pilar: {proj['PILARES']}\n"
    mensagem += f"   📈 Projeção: R$ {proj['Projeção']:,.2f}\n"
    mensagem += f"   🎯 Percentual projetado: {proj['% Proj da Meta']:.2%}\n\n"

    mensagem += "📊 ANÁLISE GERAL POR PILAR:\n\n"

    for _, linha in df_ctrl.iterrows():
        pilar = linha['PILARES']
        real = linha['REAL']
        meta = linha['META']
        percentual = linha['% Real Meta']

        if pd.isna(real) or pd.isna(meta) or pd.isna(percentual):
            continue

        if percentual >= 1:
            mensagem += f"✅ {pilar}: Dentro da meta (R$ {real:,.2f} de R$ {meta:,.2f} | {percentual:.2%})\n"
        else:
            deficit = meta - real
            mensagem += f"⚠️ {pilar}: Abaixo da meta em R$ {deficit:,.2f} (Real: R$ {real:,.2f} | Meta: R$ {meta:,.2f} | {percentual:.2%})\n"

    mensagem += "\n📬 Relatório enviado automaticamente."

    # === Configuração do email ===
    remetente = 'apollolds2@gmail.com'
    senha = 'bmxfpgaqezypyvon'
    destinatario = 'apollolopeeees@gmail.com'
    
    # Enviar e-mail
    yag = yagmail.SMTP(remetente, senha)
    yag.send(
        to=destinatario,
        subject='Análise de Tendências - Relatório Automático',
        contents=mensagem
    )

    print("📧 E-mail enviado com sucesso!")
