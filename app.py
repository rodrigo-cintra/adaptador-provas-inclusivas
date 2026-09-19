import ast
import io
import json
import re
import zipfile
import difflib
import requests
import streamlit as st
from docx import Document

st.set_page_config(
    page_title="Adaptador Acadêmico Inclusivo",
    page_icon="🎓",
    layout="centered"
)

st.title("🎓 Adaptação Didática Inclusiva de Avaliações")
st.markdown("""
Carregue a avaliação em formato **.docx**. O sistema processará a matriz 
cognitiva, gerará os cadernos adaptados, os gabaritos orientados e o guia de aplicação em um único pacote consolidado.
""")

DIFY_API_KEY = "app-9NqVkZLWEQgSjy2AZHZ5KGO3"
DIFY_WORKFLOW_URL = "https://api.dify.ai/v1/workflows/run"
DIFY_UPLOAD_URL = "https://api.dify.ai/v1/files/upload"

if "pacote_zip" not in st.session_state:
    st.session_state.pacote_zip = None
if "resumo_geracao" not in st.session_state:
    st.session_state.resumo_geracao = []
if "detalhes_log" not in st.session_state:
    st.session_state.detalhes_log = []

arquivo_upload = st.file_uploader("Selecione o arquivo da Prova Regular (.docx):", type=["docx"])

contexto_turma_padrao = """[
  {"perfil_id": "TDAH_01", "alunos": ["Lucas Silva", "Gabriel Santos"]},
  {"perfil_id": "TEA_SUPORTE1", "alunos": ["Beatriz Mendes"]}
]"""

contexto_turma = st.text_area(
    "Mapeamento de Perfis da Turma (JSON):",
    value=contexto_turma_padrao,
    height=120
)

def limpar_string(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r'[\r\n\t]+', ' ', str(s))
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def preencher_paragrafo_com_markdown(paragrafo, texto_formatado: str):
    """Substitui o conteúdo do parágrafo aplicando negrito real para padrões **texto**."""
    paragrafo.text = ""
    partes = re.split(r'(\*\*.*?\*\*)', texto_formatado)
    for parte in partes:
        if parte.startswith('**') and parte.endswith('**') and len(parte) >= 4:
            run = paragrafo.add_run(parte[2:-2])
            run.bold = True
        else:
            paragrafo.add_run(parte)

def substituir_em_paragrafo(paragrafo, texto_antigo: str, texto_novo: str) -> bool:
    texto_p = paragrafo.text
    if not texto_p.strip() or not texto_antigo.strip():
        return False

    antigo_limpo = limpar_string(texto_antigo)
    p_limpo = limpar_string(texto_p)

    # 1. Correspondência exata direta
    if texto_antigo in texto_p:
        preencher_paragrafo_com_markdown(paragrafo, texto_p.replace(texto_antigo, texto_novo))
        return True

    # 2. Correspondência normalizada sem pontuações ou espaços residuais
    if antigo_limpo in p_limpo or p_limpo in antigo_limpo:
        if len(antigo_limpo) >= 15:
            preencher_paragrafo_com_markdown(paragrafo, texto_novo)
            return True

    # 3. Correspondência por similaridade difusa
    if len(antigo_limpo) > 20 and len(p_limpo) > 20:
        razao = difflib.SequenceMatcher(None, antigo_limpo, p_limpo).ratio()
        if razao >= 0.70:
            preencher_paragrafo_com_markdown(paragrafo, texto_novo)
            return True

    return False

def extrair_pares_seguro(bloco_bruto):
    """Decodifica com suporte tanto a JSON rigoroso quanto a literais com aspas simples."""
    dados = None
    
    if isinstance(bloco_bruto, dict) and "conteudo" in bloco_bruto:
        bloco_bruto = bloco_bruto["conteudo"]

    if isinstance(bloco_bruto, str):
        texto = bloco_bruto.strip()
        if texto.startswith("```json"):
            texto = texto[7:]
        if texto.startswith("```"):
            texto = texto[3:]
        if texto.endswith("```"):
            texto = texto[:-3]
        texto = texto.strip()

        try:
            dados = json.loads(texto)
        except Exception:
            try:
                dados = ast.literal_eval(texto)
            except Exception:
                dados = []
    elif isinstance(bloco_bruto, (list, dict)):
        dados = bloco_bruto

    if isinstance(dados, dict):
        for k in ["adaptacoes", "questoes_adaptadas", "questoes", "conteudo", "result"]:
            if k in dados and isinstance(dados[k], (list, str)):
                if isinstance(dados[k], str):
                    return extrair_pares_seguro(dados[k])
                dados = dados[k]
                break
        if isinstance(dados, dict):
            dados = [dados]

    pares = []
    if isinstance(dados, list):
        for elem in dados:
            if not isinstance(elem, dict):
                continue
            orig = elem.get("texto_original") or elem.get("enunciado_original") or elem.get("original") or ""
            adapt = elem.get("texto_adaptado") or elem.get("enunciado_adaptado") or elem.get("adaptado") or ""
            
            orig_s = str(orig).strip()
            adapt_s = str(adapt).strip()
            if orig_s and adapt_s:
                pares.append({"original": orig_s, "adaptado": adapt_s})

    return pares

def aplicar_adaptacoes_docx(bytes_docx_original, lista_pares: list):
    doc = Document(io.BytesIO(bytes_docx_original))
    total_substituicoes = 0
    relatorio = []

    for par in lista_pares:
        original = par["original"]
        adaptado = par["adaptado"]
        substituido = False

        for p in doc.paragraphs:
            if substituir_em_paragrafo(p, original, adaptado):
                substituido = True
                total_substituicoes += 1
                relatorio.append(f"✅ Substituído: '{original[:45]}...'")
                break

        if not substituido:
            for tabela in doc.tables:
                for linha in tabela.rows:
                    for celula in linha.cells:
                        for p in celula.paragraphs:
                            if substituir_em_paragrafo(p, original, adaptado):
                                substituido = True
                                total_substituicoes += 1
                                relatorio.append(f"✅ Substituído em tabela: '{original[:45]}...'")
                                break
                        if substituido:
                            break
                    if substituido:
                        break
                if substituido:
                    break

        if not substituido:
            relatorio.append(f"❌ Não localizado no documento: '{original[:45]}...'")

    buffer_saida = io.BytesIO()
    doc.save(buffer_saida)
    buffer_saida.seek(0)
    return buffer_saida, total_substituicoes, relatorio

def gerar_documento_texto(titulo: str, secoes: dict) -> io.BytesIO:
    doc = Document()
    doc.add_heading(titulo, level=1)
    
    for subtitulo, conteudo in secoes.items():
        doc.add_heading(subtitulo, level=2)
        if isinstance(conteudo, list):
            for item in conteudo:
                doc.add_paragraph(str(item), style='List Bullet')
        elif isinstance(conteudo, dict):
            for k, v in conteudo.items():
                p = doc.add_paragraph()
                p.add_run(f"{k}: ").bold = True
                p.add_run(str(v))
        else:
            doc.add_paragraph(str(conteudo))
            
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

if st.button("Gerar Pacote Pedagógico Completo", type="primary"):
    if not arquivo_upload:
        st.warning("Por favor, selecione um arquivo .docx antes de prosseguir.")
    else:
        st.session_state.pacote_zip = None
        st.session_state.resumo_geracao = []
        st.session_state.detalhes_log = []

        with st.status("Processando pacote pedagógico no Dify...", expanded=True) as status:
            try:
                bytes_docx = arquivo_upload.read()
                headers_auth = {"Authorization": f"Bearer {DIFY_API_KEY}"}

                st.write("📤 Enviando documento institucional...")
                files = {
                    'file': (arquivo_upload.name, io.BytesIO(bytes_docx), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                }
                resp_upload = requests.post(
                    DIFY_UPLOAD_URL, 
                    headers=headers_auth, 
                    files=files, 
                    data={'user': 'docente-web'}, 
                    timeout=60
                )
                
                if resp_upload.status_code not in [200, 201]:
                    status.update(label="Erro no upload", state="error")
                    st.error(f"Erro no upload: {resp_upload.text}")
                else:
                    file_id = resp_upload.json().get("id")
                    st.write("🧠 Adaptando questões, gabaritos e critérios inclusivos...")

                    payload = {
                        "inputs": {
                            "arquivo_prova": [
                                {
                                    "type": "document",
                                    "transfer_method": "local_file",
                                    "upload_file_id": file_id
                                }
                            ],
                            "contexto_turma": contexto_turma
                        },
                        "response_mode": "streaming",
                        "user": "docente-web"
                    }

                    resposta = requests.post(
                        DIFY_WORKFLOW_URL, 
                        headers={**headers_auth, "Content-Type": "application/json"}, 
                        json=payload,
                        stream=True,
                        timeout=600
                    )

                    outputs_finais = None
                    erro_fluxo = None
                    
                    for linha in resposta.iter_lines():
                        if linha:
                            linha_str = linha.decode('utf-8')
                            if linha_str.startswith("data:"):
                                corpo = linha_str[5:].strip()
                                if corpo:
                                    try:
                                        dados_evento = json.loads(corpo)
                                        evento = dados_evento.get("event
