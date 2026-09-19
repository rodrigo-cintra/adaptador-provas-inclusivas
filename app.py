import io
import json
import re
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
cognitiva e gerará os cadernos adaptados mantendo rigorosamente a identidade visual original.
""")

# Configurações de ligação ao Dify
DIFY_API_KEY = "app-9NqVkZLWEQgSjy2AZHZ5KGO3"
DIFY_WORKFLOW_URL = "https://api.dify.ai/v1/workflows/run"
DIFY_UPLOAD_URL = "https://api.dify.ai/v1/files/upload"

# Upload do ficheiro da avaliação
arquivo_upload = st.file_uploader("Selecione o arquivo da Prova (.docx):", type=["docx"])

contexto_turma_padrao = """[
  {"perfil_id": "TDAH_01", "alunos": ["Lucas Silva", "Gabriel Santos"]},
  {"perfil_id": "TEA_SUPORTE1", "alunos": ["Beatriz Mendes"]}
]"""

contexto_turma = st.text_area(
    "Mapeamento de Perfis da Turma (JSON):",
    value=contexto_turma_padrao,
    height=120
)

def limpar_espacos(texto: str) -> str:
    return re.sub(r'\s+', ' ', texto).strip()

def substituir_no_paragrafo(paragrafo, texto_antigo: str, texto_novo: str) -> bool:
    texto_p = paragrafo.text
    if texto_antigo in texto_p:
        if paragrafo.runs:
            primeiro = paragrafo.runs[0]
            novo = texto_p.replace(texto_antigo, texto_novo)
            primeiro.text = novo
            for r in paragrafo.runs[1:]:
                r.text = ""
        else:
            paragrafo.text = texto_p.replace(texto_antigo, texto_novo)
        return True
    return False

def aplicar_adaptacoes_docx(bytes_docx_original, lista_adaptacoes: list) -> io.BytesIO:
    doc = Document(io.BytesIO(bytes_docx_original))
    
    for item in lista_adaptacoes:
        original = item.get("texto_original", "").strip()
        adaptado = item.get("texto_adaptado", "").strip()
        
        if not original or not adaptado:
            continue

        substituido = False
        # Varredura nos parágrafos principais
        for p in doc.paragraphs:
            if original in p.text or limpar_espacos(original) in limpar_espacos(p.text):
                substituir_no_paragrafo(p, original, adaptado)
                substituido = True
                break

        # Varredura em células de tabelas
        if not substituido:
            for tabela in doc.tables:
                for linha in tabela.rows:
                    for celula in linha.cells:
                        for p in celula.paragraphs:
                            if original in p.text or limpar_espacos(original) in limpar_espacos(p.text):
                                substituir_no_paragrafo(p, original, adaptado)
                                substituido = True
                                break
                        if substituido:
                            break
                    if substituido:
                        break
                if substituido:
                    break

    buffer_saida = io.BytesIO()
    doc.save(buffer_saida)
    buffer_saida.seek(0)
    return buffer_saida

if st.button("Gerar Avaliações Adaptadas", type="primary"):
    if not arquivo_upload:
        st.warning("Por favor, selecione um arquivo .docx antes de prosseguir.")
    else:
        with st.spinner("Enviando arquivo e processando adaptações no Dify..."):
            try:
                bytes_docx = arquivo_upload.read()

                headers_auth = {"Authorization": f"Bearer {DIFY_API_KEY}"}

                # 1. Carregamento do arquivo para a API do Dify
                files = {
                    'file': (arquivo_upload.name, io.BytesIO(bytes_docx), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                }
                data_upload = {'user': 'docente-web'}
                
                resp_upload = requests.post(DIFY_UPLOAD_URL, headers=headers_auth, files=files, data=data_upload)
                
                if resp_upload.status_code not in [200, 201]:
                    st.error(f"Erro no carregamento do arquivo para o Dify: {resp_upload.text}")
                else:
                    file_id = resp_upload.json().get("id")

                    # 2. Execução do fluxo passando arquivo_prova como LISTA de arquivos
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
                        "response_mode": "blocking",
                        "user": "docente-web"
                    }

                    resposta = requests.post(
                        DIFY_WORKFLOW_URL, 
                        headers={**headers_auth, "Content-Type": "application/json"}, 
                        json=payload
                    )

                    if resposta.status_code != 200:
                        st.error(f"Erro ao consultar o Dify: {resposta.text}")
                    else:
                        dados_saida = resposta.json().get("data", {}).get("outputs", {})
                        resultado_perfis = dados_saida.get("resultado_perfis", [])

                        if not resultado_perfis:
                            st.warning("Nenhum dado retornado pelo workflow do Dify.")
                        else:
                            st.success("Adaptação concluída com sucesso!")
                            
                            for idx, bloco in enumerate(resultado_perfis):
                                lista_itens = json.loads(bloco) if isinstance(bloco, str) else bloco
                                
                                arquivo_modificado = aplicar_adaptacoes_docx(bytes_docx, lista_itens)
                                nome_saida = f"avaliacao_adaptada_perfil_{idx + 1}.docx"

                                st.download_button(
                                    label=f"📥 Baixar Caderno Adaptado #{idx + 1} (.docx)",
                                    data=arquivo_modificado,
                                    file_name=nome_saida,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key=f"btn_dl_{idx}"
                                )

            except Exception as e:
                st.error(f"Ocorreu um erro no processamento: {str(e)}")
