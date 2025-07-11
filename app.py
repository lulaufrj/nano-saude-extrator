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

def identificar_blocos(linhas):
    # Encontrar início do título (primeira linha longa)
    idx_titulo = next((ix for ix, l in enumerate(linhas) if len(l.strip()) > 10), None)
    if idx_titulo is None:
        return '', '', ''
    # Pega o bloco do título: vai até linha que claramente contenha nomes próprios ou asterisco (autor)
    bloco_titulo = []
    for ix, l in enumerate(linhas[idx_titulo:], start=idx_titulo):
        if re.search(r'([A-Z][a-z]+ [A-Z][\w\.-]+.*[,\*])', l) or (',' in l and '*' in l):
            idx_autor = ix
            break
        bloco_titulo.append(l)
    else:
        idx_autor = idx_titulo + len(bloco_titulo)
    titulo = ' '.join(bloco_titulo)
    # Pega bloco de autores: linhas curtas sem palavras-chave institucionais
    bloco_autores = []
    idx_fim_autor = idx_autor
    for ix, l in enumerate(linhas[idx_autor:], start=idx_autor):
        if (re.search(r'(institut|universid|@|palavras|keywords|introduction|abstract|\b[Cc]orresponding|[Ee]mail|\b[0-9]{1,2}\.|\b[1-9]\))', l) or len(l.split()) > 20):
            break
        bloco_autores.append(l)
        idx_fim_autor = ix + 1
    autores_texto = ' '.join(bloco_autores)
    return titulo, autores_texto, idx_fim_autor

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
        titulo, autores_texto, _ = identificar_blocos(linhas)
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
