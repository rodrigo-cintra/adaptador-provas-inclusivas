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

# Inicialização do estado de sessão para persistência
if "cadernos_gerados" not in st.session_state:
    st.session_state.cadernos_gerados = []
if "log_substituicoes" not in st.session_state:
    st.session_state.log_substituicoes = []

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

def normalizar_texto(texto: str) -> str:
    """Normaliza espaços em branco e quebras de linha para comparação confiável."""
    if not texto:
        return ""
    return re.sub(r'\s+', ' ', texto).strip()

def substituir_no_paragrafo(paragrafo, texto_antigo: str, texto_novo: str) -> bool:
    texto_p = paragrafo.text
    alvo = normalizar_texto(texto_antigo)
    atual_norm = normalizar_texto(texto_p)

    if texto_antigo in texto_p:
        if paragrafo.runs:
            primeiro = paragrafo.runs[0]
            primeiro.text = texto_p.replace(texto_antigo, texto_novo)
            for r in paragrafo.runs[1:]:
                r.text = ""
        else:
            paragrafo.text = texto_p.replace(texto_antigo, texto_novo)
        return True
    elif alvo and alvo in atual_norm:
        # Se coincidir de forma normalizada, substitui preservando o primeiro run
        if paragrafo.runs:
            paragrafo.runs[0].text = texto_novo
            for r in paragrafo.runs[1:]:
                r.text = ""
        else:
            paragrafo.text = texto_novo
        return True
    return False

def extrair_lista_adaptacoes(bloco_bruto):
    """Garante a conversão de strings JSON ou dicionários em listas de adaptações."""
    if isinstance(bloco_bruto, str):
        try:
            limpo = bloco_bruto.strip()
            # Remove cercaduras markdown se houver
            if limpo.startswith("```json"):
                limpo = limpo[7:]
            if limpo.startswith("```"):
                limpo = limpo[3:]
            if limpo.endswith("```"):
                limpo = limpo[:-3]
            bloco_bruto = json.loads(limpo.strip())
        except Exception:
            return []

    if isinstance(bloco_bruto, dict):
        for chave in ["adaptacoes", "questoes_adaptadas", "itens", "output", "result"]:
            if chave in bloco_bruto and isinstance(bloco_bruto[chave], list):
                return bloco_bruto[chave]
        return [bloco_bruto]
    elif isinstance(bloco_bruto, list):
        return bloco_bruto
    return []

def aplicar_adaptacoes_docx(bytes_docx_original, lista_adaptacoes: list):
    doc = Document(io.BytesIO(bytes_docx_original))
    total_substituicoes = 0
    relatorio = []

    for item in lista_adaptacoes:
        if not isinstance(item, dict):
            continue

        # Mapeamento flexível das chaves geradas pelo LLM
        original = item.get("texto_original") or item.get("enunciado_original") or item.get("original") or ""
        adaptado = item.get("texto_adaptado") or item.get("enunciado_adaptado") or item.get("adaptado") or ""

        original = str(original).strip()
        adaptado = str(adaptado).strip()

        if not original or not adaptado:
            continue

        substituido = False

        # Varredura nos parágrafos principais
        for p in doc.paragraphs:
            if substituir_no_paragrafo(p, original, adaptado):
                substituido = True
                total_substituicoes += 1
                relatorio.append(f"Substituído com sucesso: '{original[:40]}...'")
                break

        # Varredura em células de tabelas
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
            relatorio.append(f"Não localizado no documento: '{original[:40]}...'")

    buffer_saida = io.BytesIO()
    doc.save(buffer_saida)
    buffer_saida.seek(0)
    return buffer_saida, total_substituicoes, relatorio

# Disparo do processamento
if st.button("Gerar Avaliações Adaptadas", type="primary"):
    if not arquivo_upload:
        st.warning("Por favor, selecione um arquivo .docx antes de prosseguir.")
    else:
        # Limpa execuções anteriores
        st.session_state.cadernos_gerados = []
        st.session_state.log_substituicoes = []

        with st.status("Processando avaliação no Dify Cloud...", expanded=True) as status:
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
                    status.update(label="Erro no envio do arquivo", state="error")
                    st.error(f"Erro no upload: {resp_upload.text}")
                else:
                    file_id = resp_upload.json().get("id")
                    st.write("🧠 Processando taxonomia pedagógica e perfis...")

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
                                                st.write(f"⚙️ Executando: {no_nome}...")
                                        elif evento == "node_finished":
                                            dados_no = dados_evento.get("data", {})
                                            if dados_no.get("status") == "failed":
                                                erro_fluxo = f"Falha no nó '{dados_no.get('title')}': {dados_no.get('error')}"
                                    except Exception:
                                        pass

                    if erro_fluxo:
                        status.update(label="Falha no Workflow", state="error")
                        st.error(f"Detalhes do erro no Dify: {erro_fluxo}")
                    elif not outputs_finais:
                        status.update(label="Erro na execução", state="error")
                        st.error("O fluxo encerrou sem entregar as saídas configuradas.")
                    else:
                        resultado_perfis = outputs_finais.get("resultado_perfis", [])
                        
                        if not resultado_perfis:
                            status.update(label="Concluído com avisos", state="complete")
                            st.warning("A saída 'resultado_perfis' veio vazia do Dify.")
                        else:
                            # Processa e persiste os arquivos gerados na sessão
                            for idx, bloco in enumerate(resultado_perfis):
                                lista_itens = extrair_lista_adaptacoes(bloco)
                                docx_adaptado, total_subs, relatorio = aplicar_adaptacoes_docx(bytes_docx, lista_itens)
                                
                                nome_saida = f"avaliacao_adaptada_perfil_{idx + 1}.docx"
                                st.session_state.cadernos_gerados.append({
                                    "nome": nome_saida,
                                    "bytes": docx_adaptado.getvalue(),
                                    "indice": idx + 1,
                                    "alteracoes": total_subs
                                })
                                st.session_state.log_substituicoes.append({
                                    "perfil": idx + 1,
                                    "detalhes": relatorio,
                                    "raw_output": lista_itens
                                })

                            status.update(label="Cadernos gerados com sucesso!", state="complete")

            except Exception as e:
                status.update(label="Erro inesperado", state="error")
                st.error(f"Falha durante a execução: {str(e)}")

# Exibição persistente dos botões de download e auditoria
if st.session_state.cadernos_gerados:
    st.divider()
    st.subheader("📥 Cadernos Adaptados Disponíveis")
    
    col1, col2 = st.columns(len(st.session_state.cadernos_gerados))
    cols = [col1, col2] if len(st.session_state.cadernos_gerados) == 2 else [st]
    
    for i, caderno in enumerate(st.session_state.cadernos_gerados):
        col_atual = cols[i] if i < len(cols) else st
        with col_atual:
            st.metric(
                label=f"Caderno #{caderno['indice']}", 
                value=f"{caderno['alteracoes']} alterações"
            )
            st.download_button(
                label=f"⬇️ Baixar Caderno #{caderno['indice']} (.docx)",
                data=caderno['bytes'],
                file_name=caderno['nome'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"dl_persistente_{caderno['indice']}"
            )

    with st.expander("🔍 Ver detalhes técnicos das alterações aplicadas"):
        for log in st.session_state.log_substituicoes:
            st.markdown(f"**Perfil #{log['perfil']}**")
            for linha in log['detalhes']:
                st.text(linha)
            st.caption("JSON recebido do Dify:")
            st.json(log['raw_output'])
