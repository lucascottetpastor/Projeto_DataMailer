import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime

template_path = r"Caminho\Do\Template\Fundo"
roboto_bold_path = r"Caminho\Fonte\Roboto"
arquivo_excel = r"Caminho\Arquivo\Excel"

pdfmetrics.registerFont(TTFont("Roboto-Bold", roboto_bold_path))

def abreviar_nome(nome_completo):
    partes = nome_completo.split()
    if len(partes) > 1:
        return f"{partes[0]} {partes[1][0]}. {partes[-1]}"
    return partes[0]

def formatar_numero(x):
    return f"{x:,.0f}".replace(",", ".")

def calcular_dados(arquivo_excel):
    df = pd.read_excel(arquivo_excel)
    
    df['md_avaliacao'] = df['md_avaliacao'].astype(str).replace(',', '.', regex=True).astype(float)
    
    alunos_ativos = len(df['id_usuario'])
    total_concluidos = df['tt_cursos_concluidos'].sum()
    total_iniciados = df['tt_cursos_iniciados'].sum()
    total_cursos = df['num_cursos'].sum()
    media_total_cursos = df['num_cursos'].mean()
    total_nao_iniciados = total_cursos - (total_concluidos + total_iniciados)
    trilha = df['num_trilhas'].mean()
    trilha = f"{trilha:.1f}"
    matriculas_geradas = total_nao_iniciados + total_concluidos + total_iniciados
    percentual_cursos_concluidos = (total_concluidos / matriculas_geradas) * 100 if matriculas_geradas > 0 else 0
    percentual_cursos_nao_iniciados = (total_nao_iniciados / matriculas_geradas) * 100 if matriculas_geradas > 0 else 0
    percentual_cursos_iniciados = (total_iniciados / matriculas_geradas) * 100 if matriculas_geradas > 0 else 0

    avaliacoes_validas = df['md_avaliacao'].replace(0, pd.NA).dropna()
    avaliacoes_normalizadas = avaliacoes_validas / 10
    media_conhecimento = avaliacoes_normalizadas.mean()
    
    alunos_ativos_formatado = formatar_numero(alunos_ativos)
    matriculas_geradas_formatado = formatar_numero(matriculas_geradas)
    total_nao_iniciados_formatado = formatar_numero(total_nao_iniciados)
    total_concluidos_formatado = formatar_numero(total_concluidos)
    total_iniciados_formatado = formatar_numero(total_iniciados)

    maior_nota = df['md_avaliacao'].max()
    candidatos_sabe_mais = df[df['md_avaliacao'] == maior_nota]
    if len(candidatos_sabe_mais) > 1:
        quem_sabe_mais = candidatos_sabe_mais.loc[candidatos_sabe_mais['tt_cursos_concluidos'].idxmax()]
    else:
        quem_sabe_mais = candidatos_sabe_mais.iloc[0]
    id_quem_sabe_mais = quem_sabe_mais['id_usuario']
    nome_quem_sabe_mais = abreviar_nome(quem_sabe_mais['nm_usuario'])
    perfil_quem_sabe_mais = "Cod. Usuário " + str(id_quem_sabe_mais)
    media_quem_sabe_mais = quem_sabe_mais['md_avaliacao']
    
    maior_treino = df['tt_cursos_concluidos'].max()
    candidatos_treina_mais = df[df['tt_cursos_concluidos'] == maior_treino]
    if len(candidatos_treina_mais) > 1:
        quem_treina_mais = candidatos_treina_mais.loc[candidatos_treina_mais['md_avaliacao'].idxmax()]
    else:
        quem_treina_mais = candidatos_treina_mais.iloc[0]
    id_quem_treina_mais = quem_treina_mais['id_usuario']
    nome_quem_treina_mais = abreviar_nome(quem_treina_mais['nm_usuario'])
    perfil_quem_treina_mais = "Cod. Usuário " + str(id_quem_treina_mais)
    progresso_quem_treina_mais = f"{quem_treina_mais['tt_cursos_concluidos']} cursos concluídos"
    
    nome_projeto = df['empresa'].iloc[0]
    data_execucao = datetime.now().strftime("%d/%m/%Y")

    dados_calculados = {
        "alunos_ativos": alunos_ativos_formatado,
        "trilhas": str(trilha),
        "cursos": f"{media_total_cursos:.1f}",
        "matriculas_geradas": matriculas_geradas_formatado,
        "cursos_concluidos": f"{percentual_cursos_concluidos:.1f}%",
        "cursos_concluidos_secundario": total_concluidos_formatado,
        "cursos_nao_iniciados": f"{percentual_cursos_nao_iniciados:.1f}%",
        "cursos_nao_iniciados_secundario": total_nao_iniciados_formatado,
        "cursos_iniciados": f"{percentual_cursos_iniciados:.1f}%",
        "cursos_iniciados_secundario": total_iniciados_formatado,
        "media_conhecimento": f"{media_conhecimento:.1f}",
        "quem_treina_mais": str(nome_quem_treina_mais),
        "quem_treina_mais_titulo": perfil_quem_treina_mais,
        "quem_treina_mais_secundario": progresso_quem_treina_mais,
        "quem_sabe_mais": str(nome_quem_sabe_mais),
        "quem_sabe_mais_titulo": perfil_quem_sabe_mais,
        "quem_sabe_mais_secundario": f"{media_quem_sabe_mais / 10:.1f} nas avaliações",
        "nome_projeto": nome_projeto,
        "data_execucao": f"Dados até {data_execucao}",
    }
    
    return dados_calculados

dados = calcular_dados(arquivo_excel)

configuracoes = {
    "alunos_ativos": {"posicao": (98, 640), "fonte": "Roboto-Bold", "tamanho": 33, "cor": (0, 0, 0)},
    "trilhas": {"posicao": (419, 640), "fonte": "Roboto-Bold", "tamanho": 33, "cor": (0, 0, 0)},
    "cursos": {"posicao": (350, 560), "fonte": "Roboto-Bold", "tamanho": 33, "cor": (0, 0, 0)},
    "matriculas_geradas": {"posicao": (151, 470), "fonte": "Roboto-Bold", "tamanho": 33, "cor": (0, 0, 0)},
    "cursos_concluidos": {"posicao": (129, 287), "fonte": "Roboto-Bold", "tamanho": 33, "cor": (0, 0, 0)},
    "cursos_concluidos_secundario": {"posicao": (129, 272), "fonte": "Roboto-Bold", "tamanho": 15, "cor": (0, 0, 0)},
    "cursos_nao_iniciados": {"posicao": (380, 420), "fonte": "Roboto-Bold", "tamanho": 32, "cor": (0, 0, 0)},
    "cursos_nao_iniciados_secundario": {"posicao": (384, 405), "fonte": "Roboto-Bold", "tamanho": 15, "cor": (0, 0, 0)},
    "cursos_iniciados": {"posicao": (380, 295), "fonte": "Roboto-Bold", "tamanho": 32, "cor": (0, 0, 0)},
    "cursos_iniciados_secundario": {"posicao": (384, 280), "fonte": "Roboto-Bold", "tamanho": 15, "cor": (0, 0, 0)},
    "media_conhecimento": {"posicao": (411, 167), "fonte": "Roboto-Bold", "tamanho": 33, "cor": (0, 0, 0)},
    "quem_treina_mais": {"posicao": (255, 115), "fonte": "Roboto-Bold", "tamanho": 10, "cor": (0, 0, 0)},
    "quem_treina_mais_titulo": {"posicao": (255, 105), "fonte": "Roboto-Bold", "tamanho": 7, "cor": (0, 0, 0)},
    "quem_treina_mais_secundario": {"posicao": (255, 95), "fonte": "Roboto-Bold", "tamanho": 7, "cor": (0, 0, 0)},
    "quem_sabe_mais": {"posicao": (255, 75), "fonte": "Roboto-Bold", "tamanho": 10, "cor": (0, 0, 0)},
    "quem_sabe_mais_titulo": {"posicao": (255, 65), "fonte": "Roboto-Bold", "tamanho": 7, "cor": (0, 0, 0)},
    "quem_sabe_mais_secundario": {"posicao": (255, 55), "fonte": "Roboto-Bold", "tamanho": 7, "cor": (0, 0, 0)},
    "nome_projeto": {"posicao": (150, 708), "fonte": "Roboto-Bold", "tamanho": 10, "cor": (0, 0, 0)},
    "data_execucao": {"posicao": (362, 708), "fonte": "Roboto-Bold", "tamanho": 10, "cor": (0, 0, 0)},
}

def desenhar_fundo_e_dados(canvas):
    canvas.drawImage(template_path, 0, 0, width=letter[0], height=letter[1], preserveAspectRatio=True, mask='auto')
    for chave, config in configuracoes.items():
        x, y = config["posicao"]
        canvas.setFont(config["fonte"], config["tamanho"])
        canvas.setFillColorRGB(*config["cor"])
        canvas.drawString(x, y, dados[chave])

def gerar_pdf_com_template():
    pdf_buffer = io.BytesIO()
    c = pdf_canvas.Canvas(pdf_buffer, pagesize=letter)
    
    c.setTitle(f"{nome_arquivo_anexo}")
    desenhar_fundo_e_dados(c)
    
    c.showPage()
    c.save()
    
    pdf_buffer.seek(0)
    return pdf_buffer


nome_cliente = dados["nome_projeto"]
user_exibicao = 'NOME DO USUARIO'
user = 'E-MAIL PARA ENVIO'
password = 'SENHA DO E-MAIL PARA ENVIO' #Senha de APP
destinatario = 'E-MAIL DO DESTINATÁRIO'
assunto = f'ASSUNTO DO E-MAIL'
corpo = f''' 
<!DOCTYPE html>
<html>
    <p>Olá!</p> 
    <p>Espero que esteja bem.</p> 
    
    <p> Desenvolvido por Lucas Cottet Pastor </p>
</html>'''
data_nome_anexo = datetime.now().strftime("%d.%m.%Y")
nome_arquivo_anexo = f'Resumo {data_nome_anexo}.pdf'

def enviar_email(destinatario, assunto, corpo, anexo_bytes, nome_arquivo_anexo, user, password):
    msg = MIMEMultipart()
    msg['From'] = f'{user_exibicao } <{user}>'
    msg['To'] = destinatario
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo, 'html'))

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(anexo_bytes.getvalue())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={nome_arquivo_anexo}')
    msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587) #SMTP para Gmail.
    server.set_debuglevel(0)
    server.login(user, password)
    server.sendmail(user, destinatario, msg.as_string())
    server.quit()
    print(f"Email enviado para {destinatario}")

pdf_buffer = gerar_pdf_com_template()
enviar_email(
    destinatario=destinatario,
    assunto=assunto,
    corpo=corpo,
    anexo_bytes=pdf_buffer,
    nome_arquivo_anexo=nome_arquivo_anexo,
    user=user,
    password=password
)
