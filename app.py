import ast
import io
import json
import re
import zipfile
import difflib
import requests
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

st.set_page_config(
    page_title="Adaptador Acadêmico Inclusivo",
    page_icon="🎓",
    layout="centered"
)

st.title("🎓 Adaptação Didática Inclusiva de Avaliações")
st.markdown("""
Carregue a avaliação em formato **.docx**. O sistema processará a matriz 
cognitiva, gerará os cadernos adaptados, as rubricas analíticas aprofundadas e o guia de aplicação em um único pacote consolidado.
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

    if texto_antigo in texto_p:
        preencher_paragrafo_com_markdown(paragrafo, texto_p.replace(texto_antigo, texto_novo))
        return True

    if antigo_limpo in p_limpo or p_limpo in antigo_limpo:
        if len(antigo_limpo) >= 15:
            preencher_paragrafo_com_markdown(paragrafo, texto_novo)
            return True

    if len(antigo_limpo) > 20 and len(p_limpo) > 20:
        razao = difflib.SequenceMatcher(None, antigo_limpo, p_limpo).ratio()
        if razao >= 0.70:
            preencher_paragrafo_com_markdown(paragrafo, texto_novo)
            return True

    return False

def extrair_dados_perfil(bloco_bruto):
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

    itens_adaptados = []
    if isinstance(dados, list):
        itens_adaptados = dados
    elif isinstance(dados, dict):
        for k in ["adaptacoes", "questoes_adaptadas", "questoes", "conteudo", "result"]:
            if k in dados and isinstance(dados[k], list):
                itens_adaptados = dados[k]
                break
        if not itens_adaptados and ("texto_original" in dados or "original" in dados):
            itens_adaptados = [dados]

    pares = []
    for elem in itens_adaptados:
        if not isinstance(elem, dict):
            continue
        orig = elem.get("texto_original") or elem.get("enunciado_original") or elem.get("original") or ""
        adapt = elem.get("texto_adaptado") or elem.get("enunciado_adaptado") or elem.get("adaptado") or ""
        num = elem.get("numero_item") or elem.get("item") or len(pares) + 1
        
        orig_s = str(orig).strip()
        adapt_s = str(adapt).strip()
        if orig_s and adapt_s:
            pares.append({
                "numero": num,
                "original": orig_s,
                "adaptado": adapt_s,
                "dados_completos": elem
            })
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
                relatorio.append(f"Substituído: {original[:40]}...")
                break

        if not substituido:
            for tabela in doc.tables:
                for linha in tabela.rows:
                    for celula in linha.cells:
                        for p in celula.paragraphs:
                            if substituir_em_paragrafo(p, original, adaptado):
                                substituido = True
                                total_substituicoes += 1
                                relatorio.append(f"Substituído em tabela: {original[:40]}...")
                                break
                        if substituido:
                            break
                    if substituido:
                        break
                if substituido:
                    break

        if not substituido:
            relatorio.append(f"Não localizado: {original[:40]}...")

    buffer_saida = io.BytesIO()
    doc.save(buffer_saida)
    buffer_saida.seek(0)
    return buffer_saida, total_substituicoes, relatorio

def definir_celula_fundo(celula, hex_color):
    tcPr = celula._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def inferir_metadados_psicometricos(item_par, perfil_id: str) -> dict:
    dados = item_par.get("dados_completos", {})
    texto = item_par.get("original", "").lower()
    
    # Detecção taxonômica inferida por verbos de comando se não vier no JSON
    nivel_bloom = dados.get("bloom") or dados.get("nivel_cognitivo")
    if not nivel_bloom:
        if any(v in texto for v in ["avalie", "julgue", "critique", "defenda"]):
            nivel_bloom = "Avaliar (Nível 5)"
        elif any(v in texto for v in ["analise", "compare", "diferencie", "relacione"]):
            nivel_bloom = "Analisar (Nível 4)"
        elif any(v in texto for v in ["aplique", "calcule", "resolva", "demonstre"]):
            nivel_bloom = "Aplicar (Nível 3)"
        elif any(v in texto for v in ["explique", "caracterize", "descreva", "discuta"]):
            nivel_bloom = "Compreender (Nível 2)"
        else:
            nivel_bloom = "Lembrar / Identificar (Nível 1)"

    barreira = "Sobrecarga de memória operacional e dispersão atencional decorrente de enunciado denso."
    ajuste = "Segmentação em comandos unitários com destaque visual nos verbos de ação para orientar o foco executivo."
    criterio_perfil = "Aceitar respostas sintéticas e estruturadas em tópicos. Priorizar a precisão do conceito sobre o volume textual."

    if "TEA" in perfil_id.upper():
        barreira = "Ambiguidade na interpretação de comandos múltiplos e esforço excessivo com linguagem implícita ou figurada."
        ajuste = "Linearização da ordem dos comandos, uso de termos diretos e eliminação de duplos sentidos nas premissas."
        criterio_perfil = "Valorizar respostas literais e diretas. Não exigir floreios discursivos ou inferências não explícitas no comando."

    return {
        "bloom": nivel_bloom,
        "barreira": dados.get("barreira_enfrentada") or barreira,
        "ajuste": dados.get("justificativa_acessibilidade") or ajuste,
        "criterio_perfil": dados.get("criterio_especifico") or criterio_perfil
    }

def gerar_rubrica_analitica_sofisticada(perfil_id: str, lista_pares: list) -> io.BytesIO:
    doc = Document()
    
    # Configuração de margens
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Título Principal
    p_title = doc.add_paragraph()
    run_title = p_title.add_run(f"Gabarito Orientado & Matriz de Correção Analítica")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(20)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(24, 43, 73)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_sub = doc.add_paragraph()
    run_sub = p_sub.add_run(f"Instrumento Técnico de Avaliação Diferenciada | Perfil de Referência: {perfil_id}")
    run_sub.font.name = 'Calibri'
    run_sub.font.size = Pt(11)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(80, 80, 80)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Quadro de Fundamentação Pedagógica
    doc.add_heading("1. Fundamentação Pedagógica & Princípios Avaliativos", level=1)
    p_fund = doc.add_paragraph(
        "Este documento estabelece a matriz de correção técnica para o caderno adaptado, assegurando o princípio "
        "da equivalência cognitiva preconizado pelo Desenho Universal para a Aprendizagem (DUA) e pelas diretrizes "
        "institucionais de acessibilidade acadêmica. A adaptação visa eliminar barreiras instrumentais de acesso "
        "ao enunciado, mantendo integral o construto avaliativo e o rigor conceitual esperado para o componente curricular."
    )
    p_fund.paragraph_format.line_spacing = 1.15
    p_fund.paragraph_format.space_after = Pt(14)

    # Matriz Analítica por Questão
    doc.add_heading("2. Matriz Analítica de Correção por Item", level=1)

    for idx, par in enumerate(lista_pares):
        num = par.get("numero", idx + 1)
        meta = inferir_metadados_psicometricos(par, perfil_id)

        h2 = doc.add_heading(f"Item #{num} — Análise Cognitiva & Parâmetros de Desempenho", level=2)
        h2.paragraph_format.space_before = Pt(16)
        h2.paragraph_format.space_after = Pt(6)

        # Tabela Estruturada do Item
        tabela = doc.add_table(rows=6, cols=2)
        tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
        tabela.autofit = False

        largura_rotulo = Inches(2.0)
        largura_conteudo = Inches(4.5)

        dados_tabela = [
            ("Nível Taxonômico (Bloom):", meta["bloom"]),
            ("Enunciado Original:", par["original"]),
            ("Enunciado Adaptado:", par["adaptado"]),
            ("Barreira Instrumental Mitigada:", meta["barreira"]),
            ("Intervenção Didática Aplicada:", meta["ajuste"]),
            ("Diretriz Específica para o Docente:", meta["criterio_perfil"])
        ]

        for i, (rotulo, valor) in enumerate(dados_tabela):
            linha = tabela.rows[i]
            
            c_rotulo = linha.cells[0]
            c_rotulo.width = largura_rotulo
            p_r = c_rotulo.paragraphs[0]
            r_run = p_r.add_run(rotulo)
            r_run.font.name = 'Calibri'
            r_run.font.size = Pt(10)
            r_run.font.bold = True
            r_run.font.color.rgb = RGBColor(30, 30, 30)
            definir_celula_fundo(c_rotulo, "F0F2F5")

            c_val = linha.cells[1]
            c_val.width = largura_conteudo
            p_v = c_val.paragraphs[0]
            v_run = p_v.add_run(str(valor))
            v_run.font.name = 'Calibri'
            v_run.font.size = Pt(10)
            v_run.font.color.rgb = RGBColor(40, 40, 40)

        # Espaço antes da rubrica de níveis
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

        # Tabela de Rubrica Gradual (Desempenho Pleno, Parcial, Insuficiente)
        doc.add_heading(f"Rubrica de Desempenho Gradual — Item #{num}", level=3)
        tab_rubrica = doc.add_table(rows=4, cols=3)
        tab_rubrica.alignment = WD_TABLE_ALIGNMENT.CENTER
        tab_rubrica.autofit = False

        headers = ["Nível de Conquista", "Critérios Conceituais Observáveis", "Ponderação Sugerida"]
        larguras_rubrica = [Inches(1.8), Inches(3.6), Inches(1.1)]

        for c_idx, texto_h in enumerate(headers):
            cel = tab_rubrica.rows[0].cells[c_idx]
            cel.width = larguras_rubrica[c_idx]
            p_h = cel.paragraphs[0]
            r_h = p_h.add_run(texto_h)
            r_h.font.name = 'Calibri'
            r_h.font.size = Pt(9.5)
            r_h.font.bold = True
            r_h.font.color.rgb = RGBColor(255, 255, 255)
            definir_celula_fundo(cel, "1F3864")

        niveis_data = [
            (
                "Desempenho Pleno (Excelente)",
                "Mobiliza com precisão os conceitos solicitados nos comandos segmentados. Identifica, explica ou relaciona as variáveis essenciais requeridas, sem necessidade de estrutura formal extensa.",
                "90% a 100%"
            ),
            (
                "Desempenho Parcial (Suficiente)",
                "Demonstra compreensão do núcleo central do conceito, respondendo corretamente à maioria dos subcomandos, com omissão pontual de elementos secundários ou menor articulação teórica.",
                "50% a 70%"
            ),
            (
                "Desempenho Insuficiente",
                "Apresenta equívocos conceituais substantivos, fuga ao tema proposto ou ausência de nexo entre os fenômenos solicitados na questão.",
                "0% a 30%"
            )
        ]

        cores_linhas = ["FFFFFF", "F9FAFC", "FFFFFF"]
        for r_idx, (nivel, desc, pond) in enumerate(niveis_data, start=1):
            row = tab_rubrica.rows[r_idx]
            valores = [nivel, desc, pond]
            for c_idx, val in enumerate(valores):
                cel = row.cells[c_idx]
                cel.width = larguras_rubrica[c_idx]
                p = cel.paragraphs[0]
                run = p.add_run(val)
                run.font.name = 'Calibri'
                run.font.size = Pt(9)
                if c_idx == 0:
                    run.font.bold = True
                definir_celula_fundo(cel, cores_linhas[r_idx - 1])

        doc.add_paragraph().paragraph_format.space_after = Pt(16)

    # Diretrizes de Devolutiva Formativa
    doc.add_heading("3. Diretrizes para Feedback Formativo Pós-Avaliação", level=1)
    p_feed = doc.add_paragraph(
        "1. Realizar a devolutiva pontuando especificamente os conceitos atingidos em cada subcomando.\n"
        "2. Evitar apontamentos meramente punitivos sobre concisão excessiva quando a resposta estiver conceitualmente exata.\n"
        "3. Em caso de dúvidas na interpretação da resposta escrita, oportunizar esclarecimento oral breve para verificar a consolidação do construto acadêmico."
    )
    p_feed.paragraph_format.line_spacing = 1.15

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def gerar_instrucoes_aplicacao_sofisticadas(perfil_id: str) -> io.BytesIO:
    doc = Document()
    
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    p_t = doc.add_paragraph()
    r_t = p_t.add_run(f"Protocolo de Aplicação & Mediação Avaliativa Inclusiva")
    r_t.font.name = 'Calibri'
    r_t.font.size = Pt(18)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(24, 43, 73)
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run(f"Guia Operacional para o Docente e Fiscal de Sala | Perfil: {perfil_id}")
    r_sub.font.name = 'Calibri'
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(80, 80, 80)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    doc.add_heading("1. Acomodações Espaciais e Temporais", level=1)
    p1 = doc.add_paragraph(
        "• Tempo Adicional Regulamentar: Conceder até 50% de acréscimo temporal para garantir que o processamento atencional "
        "e a escrita ocorram sem penalização motora.\n"
        "• Pausas Programadas: Autorizar pausas curtas (3 a 5 minutos) a cada 45 minutos de avaliação para autorregulação "
        "neurocognitiva, com controle do tempo fora da mesa de prova.\n"
        "• Posicionamento em Sala: Garantir assento em área com menor fluxo de pessoas (afastado de portas ou janelas movimentadas), "
        "minimizando estímulos concorrentes."
    )
    p1.paragraph_format.line_spacing = 1.15

    doc.add_heading("2. Parâmetros de Mediação Permitida (O que o fiscal PODE e NÃO PODE fazer)", level=1)
    p2 = doc.add_paragraph(
        "• Leitura Mediada dos Comandos (Permitida): Se solicitado pelo estudante, o aplicador pode ler em voz alta e ritmo pausado "
        "o enunciado da questão, estritamente como redigido no caderno adaptado.\n"
        "• Esclarecimento de Terminologia Geral (Permitida): O aplicador pode clarificar o sentido de termos de comando "
        "(ex.: 'identificar', 'relacionar'), sem fornecer exemplos do conteúdo da disciplina.\n"
        "• Indução ou Validação de Resposta (Proibida): Em hipótese alguma o aplicador deve sinalizar se a resposta esboçada está "
        "certa ou incompleta."
    )
    p2.paragraph_format.line_spacing = 1.15

    doc.add_heading("3. Recursos Materiais Autorizados", level=1)
    p3 = doc.add_paragraph(
        "• Folhas de Rascunho Ilimitadas: Permitir organização de esquemas gráficos, tópicos e mapas mentais prévios à resposta final.\n"
        "• Instrumentos de Destaque: Autorizar uso livre de marca-texto, réguas guia de leitura ou canetas de cores diferenciadas "
        "para decodificação visual dos itens."
    )
    p3.paragraph_format.line_spacing = 1.15

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

                st.write("Enviando documento institucional...")
                files = {
                    'file': (
                        arquivo_upload.name,
                        io.BytesIO(bytes_docx),
                        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    )
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
                    st.write("Adaptando questões e estruturando rubricas analíticas...")

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
                        if not linha:
                            continue
                        linha_str = linha.decode('utf-8')
                        if not linha_str.startswith("data:"):
                            continue
                        corpo = linha_str[5:].strip()
                        if not corpo:
                            continue
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
                                    st.write(f"Etapa
