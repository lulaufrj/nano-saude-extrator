import streamlit as st
import pandas as pd
import re
from docx import Document
import PyPDF2
from io import BytesIO
import tempfile
import os

st.set_page_config(page_title="Extração de Resumos - Rede NanoSaúde", layout="centered")
st.title('Extração de Resumos - 3º Workshop da Rede NanoSaúde')
st.write('Faça upload dos arquivos .docx ou .pdf dos resumos para obter a tabela consolidada.')

def limpar_nome_autor(nome):
    nome = re.sub(r'[\d\*]+', '', nome)
    return nome.strip(',; .')

def extrair_linhas_docx(fpath):
    doc = Document(fpath)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]

def extrair_linhas_pdf(fpath):
    with open(fpath, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        linhas = []
        for page in reader.pages:
            texto = page.extract_text()
            if texto:
                linhas += [l.strip() for l in texto.split('\n') if l.strip()]
    return linhas

def juntar_linhas(linhas, idx_ini):
    bloco = []
    for linha in linhas[idx_ini:]:
        if not linha:
            break
        bloco.append(linha)
    return bloco

def achar_autores_bloco(linhas, inicio):
    bloco = []
    dentro = False
    for l in linhas[inicio:]:
        if (',' in l or '*' in l) and (re.search(r'[A-Za-z]', l)):
            bloco.append(l)
            dentro = True
        elif dentro and l:
            bloco.append(l)
        else:
            if dentro:
                break
    return bloco

def processar_resumos(list_files):
    trabalhos = []
    for i, fpath in enumerate(list_files):
        nome_arq = os.path.basename(fpath)
        if nome_arq.endswith('.docx'):
            linhas = extrair_linhas_docx(fpath)
        elif nome_arq.endswith('.pdf'):
            linhas = extrair_linhas_pdf(fpath)
        else:
            continue
        # Localiza primeira linha de texto relevante (título)
        idx_titulo = next((ix for ix, l in enumerate(linhas) if len(l) > 10), None)
        if idx_titulo is None:
            continue
        # Junta linhas do título (até encontrar linha vazia ou linha só com nomes de pessoas)
        titulo_bloc = [linhas[idx_titulo]]
        for off in range(idx_titulo + 1, len(linhas)):
            l = linhas[off]
            if (re.match(r'^[A-Z][a-z]+( [A-Z][a-z]+)+', l) or '*' in l or re.search(r'\d', l)) and ',' in l:
                break
            if l and len(l) > 5:
                titulo_bloc.append(l)
            else:
                break
        titulo = ' '.join(titulo_bloc)
        # Agora localiza início da linha dos autores
        idx_autores = idx_titulo + len(titulo_bloc)
        autores_bloc = achar_autores_bloco(linhas, idx_autores)
        autores_texto = ' '.join(autores_bloc)
        autores_lista = [a.strip() for a in re.split(r',|;', autores_texto) if a.strip()]
        apresentador = None
        demais = []
        for nome in autores_lista:
            if '*' in nome:
                apresentador = limpar_nome_autor(nome)
            else:
                demais.append(limpar_nome_autor(nome))
        trabalhos.append({
            'Número': f"{i+1:02}",
            'Título': titulo,
            'Apresentador': apresentador if apresentador else (demais[0] if demais else ''),
            'Demais Autores': ', '.join([a for a in demais if a])
        })
    return pd.DataFrame(trabalhos)

uploaded_files = st.file_uploader('Envie seus arquivos (.docx, .pdf)', accept_multiple_files=True)

if uploaded_files:
    temp_files = []
    for f in uploaded_files:
        with tempfile.NamedTemporaryFile(delete=False, suffix=f.name[-5:]) as tmp:
            tmp.write(f.read())
            temp_files.append(tmp.name)
    tabela = processar_resumos(temp_files)
    st.dataframe(tabela)
    # Download Excel
    towrite = BytesIO()
    tabela.to_excel(towrite, index=False)
    towrite.seek(0)
    st.download_button('Baixar tabela Excel', data=towrite, file_name='resumos_extraidos.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    # Limpa arquivos temporários
    for f in temp_files:
        os.remove(f)
