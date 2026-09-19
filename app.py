import io
import json
import re
import zipfile
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

# Configurações do Dify
DIFY_API_KEY = "app-9NqVkZLWEQgSjy2AZHZ5KGO3"
DIFY_WORKFLOW_URL = "https://api.dify.ai/v1/workflows/run"
DIFY_UPLOAD_URL = "https://api.dify.ai/v1/files/upload"

# Gerenciamento de Estado de Sessão
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

def normalizar_texto(texto: str) -> str:
    if not texto:
        return ""
    return re.sub(r'\s+', ' ', str(texto)).strip()

def substituir_no_paragrafo(paragrafo, texto_antigo: str, texto_novo: str) -> bool:
    texto_p = paragrafo.text
    alvo = normalizar_texto(texto_antigo)
    atual_norm = normalizar_texto(texto_p)

    if texto_antigo in texto_p:
        if paragrafo.runs:
            paragrafo.runs[0].text = texto_p.replace(texto_antigo, texto_novo)
            for r in paragrafo.runs[1:]:
                r.text = ""
        else:
            paragrafo.text = texto_p.replace(texto_antigo, texto_novo)
        return True
    elif alvo and alvo in atual_norm:
        if paragrafo.runs:
            paragrafo.runs[0].text = texto_novo
            for r in paragrafo.runs[1:]:
                r.text = ""
        else:
            paragrafo.text = texto_novo
        return True
    return False

def extrair_bloco(bloco_bruto):
    if isinstance(bloco_bruto, str):
        try:
            limpo = bloco_bruto.strip()
            if limpo.startswith("```json"):
                limpo = limpo[7:]
            if limpo.startswith("```"):
                limpo = limpo[3:]
            if limpo.endswith("```"):
                limpo = limpo[:-3]
            bloco_bruto = json.loads(limpo.strip())
        except Exception:
            return {}
    return bloco_bruto if isinstance(bloco_bruto, dict) else {"conteudo": str(bloco_bruto)}

def aplicar_adaptacoes_docx(bytes_docx_original, lista_adaptacoes: list):
    doc = Document(io.BytesIO(bytes_docx_original))
    total_substituicoes = 0
    relatorio = []

    for item in lista_adaptacoes:
        if not isinstance(item, dict):
            continue

        original = item.get("texto_original") or item.get("enunciado_original") or item.get("original") or ""
        adaptado = item.get("texto_adaptado") or item.get("enunciado_adaptado") or item.get("adaptado") or ""

        original = str(original).strip()
        adaptado = str(adaptado).strip()

        if not original or not adaptado:
            continue

        substituido = False
        for p in doc.paragraphs:
            if substituir_no_paragrafo(p, original, adaptado):
                substituido = True
                total_substituicoes += 1
                relatorio.append(f"Substituído: '{original[:40]}...'")
                break

        if not substituido:
            for tabela in doc.tables:
                for linha in tabela.rows:
                    for celula in linha.cells:
                        for p in celula.paragraphs:
                            if substituir_no_paragrafo(p, original, adaptado):
                                substituido = True
                                total_substituicoes += 1
                                relatorio.append(f"Substituído em tabela: '{original[:40]}...'")
                                break
                        if substituido:
                            break
                    if substituido:
                        break
                if substituido:
                    break

        if not substituido:
            relatorio.append(f"Não localizado: '{original[:40]}...'")

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
                    status.update(label="Erro no envio do documento", state="error")
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
                                        evento = dados_evento.get("event")
                                        if evento == "workflow_finished":
                                            outputs_finais = dados_evento.get("data", {}).get("outputs", {})
                                        elif evento == "workflow_failed":
                                            erro_fluxo = dados_evento.get("data", {}).get("error") or dados_evento.get("message")
                                        elif evento == "node_started":
                                            no_nome = dados_evento.get("data", {}).get("title", "")
                                            if no_nome:
                                                st.write(f"⚙️ Em andamento: {no_nome}...")
                                        elif evento == "node_finished":
                                            dados_no = dados_evento.get("data", {})
                                            if dados_no.get("status") == "failed":
                                                erro_fluxo = f"Falha no nó '{dados_no.get('title')}': {dados_no.get('error')}"
                                    except Exception:
                                        pass

                    if erro_fluxo:
                        status.update(label="Falha no processamento", state="error")
                        st.error(f"Erro no Dify: {erro_fluxo}")
                    elif not outputs_finais:
                        status.update(label="Processamento sem saída", state="error")
                        st.error("O fluxo concluiu sem gerar os dados de saída esperados.")
                    else:
                        resultado_perfis = outputs_finais.get("resultado_perfis", [])
                        
                        if not resultado_perfis:
                            status.update(label="Sem dados gerados", state="error")
                            st.warning("A variável 'resultado_perfis' retornou vazia.")
                        else:
                            buffer_zip = io.BytesIO()
                            
                            with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zip_file:
                                for idx, item_perfil in enumerate(resultado_perfis):
                                    dados_p = extrair_bloco(item_perfil)
                                    p_id = dados_p.get("perfil_id", f"PERFIL_{idx + 1}")
                                    
                                    # 1. Caderno de Prova Adaptado (.docx)
                                    lista_questoes = dados_p.get("adaptacoes") or dados_p.get("questoes_adaptadas") or []
                                    if not lista_questoes and isinstance(dados_p, list):
                                        lista_questoes = dados_p
                                        
                                    docx_adaptado, total_subs, relatorio = aplicar_adaptacoes_docx(bytes_docx, lista_questoes)
                                    nome_prova = f"{p_id}/Caderno_Prova_Adaptada_{p_id}.docx"
                                    zip_file.writestr(nome_prova, docx_adaptado.getvalue())

                                    # 2. Gabarito e Rubrica de Correção (.docx)
                                    gabarito_data = dados_p.get("rubrica_correcao") or dados_p.get("gabarito") or {
                                        "Informacao": "Rubrica pedagogica integrada.",
                                        "Equivalencia": "Avaliar o dominio conceitual sem penalizar velocidade motora."
                                    }
                                    docx_gabarito = gerar_documento_texto(
                                        f"Gabarito e Rubrica Avaliativa - {p_id}",
                                        {"Diretrizes de Correcao": gabarito_data}
                                    )
                                    nome_gabarito = f"{p_id}/Gabarito_e_Rubrica_{p_id}.docx"
                                    zip_file.writestr(nome_gabarito, docx_gabarito.getvalue())

                                    # 3. Guia de Aplicação e Mediação (.docx)
                                    instrucoes_data = dados_p.get("instrucoes_aplicacao") or dados_p.get("guia_mediacao") or {
                                        "Acomodacoes de Tempo": "Tempo adicional de ate 50% conforme diretrizes de acessibilidade.",
                                        "Mediacao do Fiscal": "Permitir leitura em voz alta se solicitado; reduzir estimulos concorrentes.",
                                        "Recursos Permitidos": "Uso de folhas de rascunho sem limite e pausas para autorregulacao."
                                    }
                                    docx_instrucoes = gerar_documento_texto(
                                        f"Instrucoes de Aplicacao - {p_id}",
                                        {"Orientacoes de Sala": instrucoes_data}
                                    )
                                    nome_instrucoes = f"{p_id}/Instrucoes_Aplicacao_{p_id}.docx"
                                    zip_file.writestr(nome_instrucoes, docx_instrucoes.getvalue())

                                    st.session_state.resumo_geracao.append({
                                        "perfil": p_id,
                                        "alteracoes": total_subs
                                    })
                                    st.session_state.detalhes_log.append({
                                        "perfil": p_id,
                                        "log": relatorio,
                                        "bruto": dados_p
                                    })

                            buffer_zip.seek(0)
                            st.session_state.pacote_zip = buffer_zip.getvalue()
                            status.update(label="Pacote pedagógico gerado com sucesso!", state="complete")

            except Exception as e:
                status.update(label="Erro no processamento", state="error")
                st.error(f"Ocorreu um erro: {str(e)}")

# Exibição do botão consolidado e métricas
if st.session_state.pacote_zip:
    st.divider()
    st.subheader("📦 Pacote Pedagógico Pronto para Download")
    st.markdown("O arquivo compactado contém, organizados por pasta de cada perfil:")
    st.markdown("✔️ Caderno de Prova Adaptado (`.docx` com layout original)")
    st.markdown("✔️ Gabarito e Rubrica Avaliativa (`.docx`)")
    st.markdown("✔️ Guia com Instruções de Aplicação para o Fiscal/Docente (`.docx`)")

    col_btn, _ = st.columns([2, 1])
    with col_btn:
        st.download_button(
            label="📥 Baixar Pacote Completo (.zip)",
            data=st.session_state.pacote_zip,
            file_name="Avaliacoes_Adaptadas_Pacote_Completo.zip",
            mime="application/zip",
            type="primary",
            key="btn_zip_consolidado"
        )

    st.markdown("---")
    cols_metrica = st.columns(len(st.session_state.resumo_geracao))
    for i, r in enumerate(st.session_state.resumo_geracao):
        cols_metrica[i].metric(label=f"Perfil: {r['perfil']}", value=f"{r['alteracoes']} modificações")

    with st.expander("🔍 Auditoria de substituições aplicadas"):
        for d in st.session_state.detalhes_log:
            st.markdown(f"**Perfil: {d['perfil']}**")
            for linha in d['log']:
                st.text(linha)
            st.caption("JSON de retorno:")
            st.json(d['bruto'])
