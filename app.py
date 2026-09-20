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

st.set_page_config(page_title="Adaptador Acadêmico Inclusivo", page_icon="🎓", layout="centered")

st.title("🎓 Adaptação Didática Inclusiva de Avaliações")
st.markdown("""
Carregue a avaliação em formato **.docx**. Selecione a estratégia de acomodação temporal 
(tempo estendido ou redução quantitativa com equivalência de construto) para gerar 
cadernos nominais, rubricas analíticas e protocolos oficiais consolidados em arquivo `.zip`.
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

# Seletor de Estratégia de Acomodação
col_est1, col_est2 = st.columns(2)
with col_est1:
    estrategia_tempo = st.selectbox(
        "Estratégia de Acomodação de Ritmo/Tempo:",
        options=[
            "Tempo Adicional Regulamentar (+50% de duração)",
            "Mesmo Tempo de Sala com Redução Quantitativa de Itens"
        ],
        help="A redução de itens evita fadiga executiva grave em estudantes com tolerância atencional reduzida."
    )

modo_reducao = "Redução Quantitativa" in estrategia_tempo

with col_est2:
    if modo_reducao:
        itens_a_manter = st.number_input(
            "Quantidade de questões a manter na prova adaptada:",
            min_value=2,
            max_value=10,
            value=4,
            step=1,
            help="O sistema selecionará os itens nucleares de maior valor epistemológico."
        )
    else:
        st.info("Serão mantidos 100% dos itens da prova com até 50% de acréscimo temporal no protocolo.")
        itens_a_manter = 999

contexto_turma_padrao = """[
  {"perfil_id": "TDAH_01", "alunos": ["Lucas Silva", "Gabriel Santos"]},
  {"perfil_id": "TEA_SUPORTE1", "alunos": ["Beatriz Mendes"]}
]"""

contexto_turma = st.text_area("Mapeamento de Perfis da Turma (JSON):", value=contexto_turma_padrao, height=110)

def sanitizar_nome(s: str) -> str:
    return re.sub(r'[^a-zA-Z0-9_-]', '_', str(s).strip())

def limpar_str(s: str) -> str:
    if not s:
        return ""
    return re.sub(r'\s+', ' ', re.sub(r'[\r\n\t]+', ' ', str(s))).strip()

def texto_consolidado(p) -> str:
    return "".join(run.text for run in p.runs) if p.runs else (p.text or "")

def preencher_md(paragrafo, texto_formatado: str):
    paragrafo.text = ""
    partes = re.split(r'(\*\*.*?\*\*)', texto_formatado)
    for parte in partes:
        if parte.startswith('**') and parte.endswith('**') and len(parte) >= 4:
            r = paragrafo.add_run(parte[2:-2])
            r.bold = True
        else:
            paragrafo.add_run(parte)

def fragmentar_comandos(texto: str) -> list:
    partes = re.split(r'(\b[a-dA-D]\)\s+)', texto)
    if len(partes) <= 1:
        return [texto]
    resultado = []
    prefixo = ""
    for pedaco in partes:
        if re.match(r'\b[a-dA-D]\)\s+', pedaco):
            prefixo = pedaco
        elif prefixo:
            resultado.append(prefixo + pedaco)
            prefixo = ""
        elif pedaco.strip():
            resultado.append(pedaco)
    return resultado if resultado else [texto]

def substituir_em_paragrafo(p, antigo: str, novo: str) -> bool:
    txt = texto_consolidado(p)
    if not txt.strip() or not antigo.strip():
        return False
    a_limpo, p_limpo = limpar_str(antigo), limpar_str(txt)

    if antigo in txt:
        preencher_md(p, txt.replace(antigo, novo))
        return True
    if (a_limpo in p_limpo or p_limpo in a_limpo) and len(a_limpo) >= 12:
        preencher_md(p, novo)
        return True
    if len(a_limpo) > 15 and len(p_limpo) > 15:
        if difflib.SequenceMatcher(None, a_limpo, p_limpo).ratio() >= 0.62:
            preencher_md(p, novo)
            return True
    return False

def injetar_nome(doc: Document, aluno: str):
    padroes = [
        r'(Nome\s*(?:do\s*Aluno\(a\)|do\s*Estudante|do\s*Aluno|Completo)?\s*:\s*)(_{2,}|\.{2,}|\s*)',
        r'(Aluno\(a\)\s*:\s*)(_{2,}|\.{2,}|\s*)',
        r'(Discente\s*:\s*)(_{2,}|\.{2,}|\s*)',
        r'(Estudante\s*:\s*)(_{2,}|\.{2,}|\s*)'
    ]
    def tentar(p):
        t = texto_consolidado(p)
        for pad in padroes:
            if re.search(pad, t, re.IGNORECASE):
                sub = re.sub(pad, lambda m: f"{m.group(1).strip()} {aluno}", t, count=1, flags=re.IGNORECASE)
                p.text = ""
                r = p.add_run(sub)
                r.bold = True
                r.font.color.rgb = RGBColor(24, 43, 73)
                return True
        return False

    for tab in doc.tables:
        for row in tab.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if tentar(p):
                        return
    for p in doc.paragraphs[:20]:
        if tentar(p):
            return
    for s in doc.sections:
        for p in s.header.paragraphs:
            if tentar(p):
                return

    p_top = doc.paragraphs[0].insert_paragraph_before() if doc.paragraphs else doc.add_paragraph()
    r1 = p_top.add_run("Estudante: ")
    r1.bold = True
    r2 = p_top.add_run(f"{aluno}\n")
    r2.bold = True
    r2.font.color.rgb = RGBColor(24, 43, 73)

def remover_item_do_documento(doc: Document, texto_alvo: str):
    """Remove parágrafos ou trechos de itens suprimidos no modo de prova reduzida."""
    alvo_limpo = limpar_str(texto_alvo)
    if not alvo_limpo or len(alvo_limpo) < 10:
        return
    for p in list(doc.paragraphs):
        txt_p = limpar_str(texto_consolidado(p))
        if alvo_limpo in txt_p or (len(alvo_limpo) > 18 and difflib.SequenceMatcher(None, alvo_limpo, txt_p).ratio() >= 0.70):
            p._element.getparent().remove(p._element)

def extrair_pares_resiliente(bloco):
    if isinstance(bloco, dict) and "conteudo" in bloco:
        bloco = bloco["conteudo"]
    if isinstance(bloco, str):
        t = bloco.strip()
        if t.startswith("```json"):
            t = t[7:]
        if t.startswith("```"):
            t = t[3:]
        if t.endswith("```"):
            t = t[:-3]
        try:
            bloco = json.loads(t.strip())
        except Exception:
            try:
                bloco = ast.literal_eval(t.strip())
            except Exception:
                bloco = []

    lista = []
    if isinstance(bloco, list):
        lista = bloco
    elif isinstance(bloco, dict):
        for k in ["adaptacoes", "questoes_adaptadas", "questoes", "conteudo", "result"]:
            if k in bloco and isinstance(bloco[k], list):
                lista = bloco[k]
                break
        if not lista and ("texto_original" in bloco or "original" in bloco):
            lista = [bloco]

    pares = []
    for elem in lista:
        if not isinstance(elem, dict):
            continue
        orig = elem.get("texto_original") or elem.get("enunciado_original") or elem.get("original") or ""
        adapt = elem.get("texto_adaptado") or elem.get("enunciado_adaptado") or elem.get("adaptado") or ""
        num = elem.get("numero_item") or elem.get("item") or len(pares) + 1

        s_orig, s_adapt = str(orig).strip(), str(adapt).strip()
        sub_origs = fragmentar_comandos(s_orig)
        sub_adapts = fragmentar_comandos(s_adapt)

        if len(sub_origs) == len(sub_adapts) and len(sub_origs) > 1:
            for i, (so, sa) in enumerate(zip(sub_origs, sub_adapts)):
                pares.append({"numero": f"{num}.{i+1}", "original": so.strip(), "adaptado": sa.strip(), "raw": elem})
        elif s_orig and s_adapt:
            pares.append({"numero": num, "original": s_orig, "adaptado": s_adapt, "raw": elem})
    return pares

def aplicar_docx_customizado(bytes_docx, pares: list, aluno: str = None, pares_suprimidos: list = None):
    doc = Document(io.BytesIO(bytes_docx))
    total_subs, logs = 0, []
    if aluno:
        injetar_nome(doc, aluno)
        logs.append(f"Nome '{aluno}' inserido.")

    # 1. Se houver redução, suprime fisicamente os itens não selecionados
    if pares_suprimidos:
        for p_sup in pares_suprimidos:
            remover_item_do_documento(doc, p_sup["original"])
        logs.append(f"Redução quantitativa aplicada: {len(pares_suprimidos)} itens suprimidos para ajuste atencional.")

    # 2. Aplica as adaptações nos itens mantidos
    for par in pares:
        orig, adapt = par["original"], par["adaptado"]
        sub = False
        for p in doc.paragraphs:
            if substituir_em_paragrafo(p, orig, adapt):
                sub = True
                total_subs += 1
                logs.append(f"Substituído: {orig[:40]}...")
                break
        if not sub:
            for tab in doc.tables:
                for row in tab.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if substituir_em_paragrafo(p, orig, adapt):
                                sub = True
                                total_subs += 1
                                logs.append(f"Substituído em tabela: {orig[:40]}...")
                                break
                        if sub:
                            break
                    if sub:
                        break
                if sub:
                    break
        if not sub:
            logs.append(f"Não localizado: {orig[:40]}...")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf, total_subs, logs

def set_fundo(cel, cor_hex):
    tcPr = cel._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), cor_hex)
    tcPr.append(shd)

def meta_psico(par, pid: str) -> dict:
    raw = par.get("raw", {})
    t = par.get("original", "").lower()
    b = raw.get("bloom") or raw.get("nivel_bloom") or raw.get("nivel_cognitivo")
    if not b:
        if any(v in t for v in ["avalie", "julgue", "critique"]):
            b = "Avaliar (Nível 5)"
        elif any(v in t for v in ["analise", "compare", "relacione"]):
            b = "Analisar (Nível 4)"
        elif any(v in t for v in ["aplique", "calcule", "resolva"]):
            b = "Aplicar (Nível 3)"
        elif any(v in t for v in ["explique", "caracterize", "descreva"]):
            b = "Compreender (Nível 2)"
        else:
            b = "Lembrar / Identificar (Nível 1)"

    barr = "Sobrecarga de memória operacional decorrente de enunciado denso."
    aj = "Segmentação em comandos unitários com destaque visual nos verbos de ação."
    crit = "Aceitar respostas sintéticas em tópicos, priorizando o rigor do conceito."
    if "TEA" in str(pid).upper():
        barr = "Ambiguidade na interpretação de comandos múltiplos e termos implícitos."
        aj = "Linearização dos comandos, vocabulário direto e eliminação de duplos sentidos."
        crit = "Valorizar respostas literais e diretas, sem exigir floreios discursivos."

    return {
        "bloom": b,
        "barreira": raw.get("barreira_enfrentada") or barr,
        "ajuste": raw.get("justificativa_acessibilidade") or aj,
        "criterio": raw.get("criterio_especifico") or crit
    }

def gerar_rubrica(pid: str, pares: list, aluno: str = None, modo_reducao: bool = False, total_orig: int = 0) -> io.BytesIO:
    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(1.0)

    p1 = doc.add_paragraph()
    r1 = p1.add_run("Gabarito Orientado & Matriz de Correção Analítica")
    r1.font.name, r1.font.size, r1.font.bold = 'Calibri', Pt(18), True
    r1.font.color.rgb = RGBColor(24, 43, 73)
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = f"Estudante: {aluno} | Perfil: {pid}" if aluno else f"Perfil Funcional: {pid}"
    p2 = doc.add_paragraph()
    r2 = p2.add_run(sub)
    r2.font.name, r2.font.size, r2.font.italic = 'Calibri', Pt(11), True
    r2.font.color.rgb = RGBColor(80, 80, 80)
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("1. Fundamentação Pedagógica & Princípios Avaliativos", level=1)
    
    texto_fund = (
        "Este documento estabelece a matriz de correção técnica para o caderno adaptado, assegurando o princípio "
        "da equivalência cognitiva preconizado pelo Desenho Universal para a Aprendizagem (DUA)."
    )
    if modo_reducao:
        texto_fund += (
            f"\n\n[PARECER DE REDUÇÃO QUANTITATIVA]: A prova regular continha {total_orig} itens e foi reestruturada "
            f"para {len(pares)} itens nucleares mantendo a mesma duração da turma. Com base na Teoria da Resposta ao Item "
            "e no controle de fadiga cognitiva executiva, a redução amostral preserva a totalidade das competências "
            "essenciais sem penalizar o estudante por déficits de sustentação atencional prolongada."
        )
    doc.add_paragraph(texto_fund)

    doc.add_heading("2. Matriz Analítica de Correção por Item", level=1)
    for idx, par in enumerate(pares):
        m = meta_psico(par, pid)
        doc.add_heading(f"Item #{par.get('numero', idx+1)} — Análise Cognitiva", level=2)
        tab = doc.add_table(rows=6, cols=2)
        tab.alignment = WD_TABLE_ALIGNMENT.CENTER
        tab.autofit = False

        dados = [
            ("Nível Bloom:", m["bloom"]),
            ("Original:", par["original"]),
            ("Adaptado:", par["adaptado"]),
            ("Barreira:", m["barreira"]),
            ("Intervenção:", m["ajuste"]),
            ("Critério Docente:", m["criterio"])
        ]
        for i, (rot, val) in enumerate(dados):
            c0, c1 = tab.rows[i].cells[0], tab.rows[i].cells[1]
            c0.width, c1.width = Inches(1.8), Inches(4.7)
            r = c0.paragraphs[0].add_run(rot)
            r.bold = True
            set_fundo(c0, "F0F2F5")
            c1.paragraphs[0].add_run(str(val))

        tab_r = doc.add_table(rows=4, cols=3)
        tab_r.alignment = WD_TABLE_ALIGNMENT.CENTER
        tab_r.autofit = False
        headers = ["Nível", "Critérios Observáveis", "Ponderação"]
        larguras = [Inches(1.8), Inches(3.6), Inches(1.1)]

        for ci, h in enumerate(headers):
            cel = tab_r.rows[0].cells[ci]
            cel.width = larguras[ci]
            r = cel.paragraphs[0].add_run(h)
            r.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            set_fundo(cel, "1F3864")

        niveis = [
            ("Pleno", "Mobiliza com precisão os conceitos solicitados nos comandos segmentados.", "90% a 100%"),
            ("Parcial", "Demonstra compreensão do núcleo central, com omissão de elementos secundários.", "50% a 70%"),
            ("Insuficiente", "Equívocos conceituais substantivos, fuga ao tema ou ausência de nexo.", "0% a 30%")
        ]
        for ri, (n, d, po) in enumerate(niveis, start=1):
            row = tab_r.rows[ri]
            for ci, v in enumerate([n, d, po]):
                c = row.cells[ci]
                c.width = larguras[ci]
                run = c.paragraphs[0].add_run(v)
                if ci == 0:
                    run.bold = True
                set_fundo(c, "FFFFFF" if ri % 2 != 0 else "F9FAFC")

    doc.add_heading("3. Diretrizes para Feedback Formativo", level=1)
    doc.add_paragraph("Pontuar conceitos atingidos e oportunizar esclarecimento oral breve em caso de concisão extrema.")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def gerar_protocolo(pid: str, aluno: str = None, modo_reducao: bool = False, total_itens_mantidos: int = 0) -> io.BytesIO:
    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(1.0)

    p1 = doc.add_paragraph()
    r1 = p1.add_run("Protocolo Oficial de Aplicação & Mediação Avaliativa")
    r1.font.name, r1.font.size, r1.font.bold = 'Calibri', Pt(18), True
    r1.font.color.rgb = RGBColor(24, 43, 73)
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = f"Estudante: {aluno} | Perfil: {pid}" if aluno else f"Diretrizes de Sala | Perfil: {pid}"
    p2 = doc.add_paragraph()
    r2 = p2.add_run(sub)
    r2.font.name, r2.font.size, r2.font.italic = 'Calibri', Pt(11), True
    r2.font.color.rgb = RGBColor(80, 80, 80)
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("1. Ficha de Parametrização & Registro de Sala", level=1)
    tab_f = doc.add_table(rows=5, cols=2)
    tab_f.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab_f.autofit = False

    tempo_desc = (
        "Mesmo tempo de sala da turma regular (Sem acréscimo temporal devido à redução de itens)" 
        if modo_reducao else 
        "[ ___ : ___ ] às [ ___ : ___ ] (com tempo estendido de até +50%)"
    )
    estrat_desc = (
        f"Redução quantitativa para {total_itens_mantidos} questões com equivalência cognitiva integral."
        if modo_reducao else
        "Manutenção integral dos itens com concessão de tempo estendido."
    )

    dados_f = [
        ("Estudante Beneficiário:", aluno if aluno else "Conforme lista homologada"),
        ("Perfil Funcional Alvo:", f"{pid} (Equivalência Cognitiva DUA)"),
        ("Estratégia Homologada:", estrat_desc),
        ("Responsável / Fiscal:", "________________________________________________________"),
        ("Duração / Horário Previsto:", tempo_desc)
    ]
    for i, (c, v) in enumerate(dados_f):
        c0, c1 = tab_f.rows[i].cells[0], tab_f.rows[i].cells[1]
        c0.width, c1.width = Inches(2.2), Inches(4.3)
        c0.paragraphs[0].add_run(c).bold = True
        set_fundo(c0, "F0F2F5")
        c1.paragraphs[0].add_run(v)

    doc.add_heading("2. Limiares de Mediação (Permitido vs. Vedado)", level=1)
    tab_m = doc.add_table(rows=3, cols=2)
    tab_m.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab_m.autofit = False

    h_med = ["Condutas Autorizadas", "Condutas Vedadas"]
    for ci, h in enumerate(h_med):
        cel = tab_m.rows[0].cells[ci]
        cel.width = Inches(3.25)
        r = cel.paragraphs[0].add_run(h)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_fundo(cel, "2E75B6" if ci == 0 else "C00000")

    regras = [
        ("Reler comandos pausadamente como impressos.", "Parafrasear conceitos ou dar pistas teóricas."),
        ("Esclarecer verbos de comando (ex: relacione).", "Validar respostas parciais durante a prova.")
    ]
    for ri, (perm, proib) in enumerate(regras, start=1):
        c0, c1 = tab_m.rows[ri].cells[0], tab_m.rows[ri].cells[1]
        c0.width, c1.width = Inches(3.25), Inches(3.25)
        c0.paragraphs[0].add_run(perm)
        set_fundo(c0, "F2F7FA")
        c1.paragraphs[0].add_run(proib)
        set_fundo(c1, "FDF2F2")

    doc.add_heading("3. Termo de Conformidade", level=1)
    doc.add_paragraph("Declaro que a avaliação foi administrada em conformidade com as diretrizes de equidade.")
    p_ass = doc.add_paragraph("\n___________________________________________________\nAssinatura do Fiscal de Sala")
    p_ass.alignment = WD_ALIGN_PARAGRAPH.CENTER

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

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
                auth = {"Authorization": f"Bearer {DIFY_API_KEY}"}

                perfis_cfg = []
                try:
                    cfg_json = json.loads(contexto_turma)
                    if isinstance(cfg_json, list):
                        perfis_cfg = cfg_json
                except Exception:
                    perfis_cfg = []

                st.write("Enviando documento institucional...")
                files = {'file': (arquivo_upload.name, io.BytesIO(bytes_docx), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')}
                resp_up = requests.post(DIFY_UPLOAD_URL, headers=auth, files=files, data={'user': 'docente-web'}, timeout=60)

                if resp_up.status_code not in [200, 201]:
                    status.update(label="Erro no upload", state="error")
                    st.error(f"Erro no upload: {resp_up.text}")
                else:
                    fid = resp_up.json().get("id")
                    st.write("Adaptando matriz taxonômica e aplicando estratégia temporal...")

                    payload = {
                        "inputs": {
                            "arquivo_prova": [{"type": "document", "transfer_method": "local_file", "upload_file_id": fid}],
                            "contexto_turma": contexto_turma
                        },
                        "response_mode": "streaming",
                        "user": "docente-web"
                    }

                    r_dify = requests.post(DIFY_WORKFLOW_URL, headers={**auth, "Content-Type": "application/json"}, json=payload, stream=True, timeout=600)

                    outputs_finais, erro_fluxo = None, None
                    for linha in r_dify.iter_lines():
                        if not linha:
                            continue
                        l_str = linha.decode('utf-8')
                        if not l_str.startswith("data:"):
                            continue
                        corpo = l_str[5:].strip()
                        if not corpo:
                            continue
                        try:
                            ev = json.loads(corpo)
                            ev_tipo = ev.get("event")
                            if ev_tipo == "workflow_finished":
                                outputs_finais = ev.get("data", {}).get("outputs", {})
                            elif ev_tipo == "workflow_failed":
                                erro_fluxo = ev.get("data", {}).get("error") or ev.get("message")
                            elif ev_tipo == "node_started":
                                n = ev.get("data", {}).get("title", "")
                                if n:
                                    st.write(f"Etapa: {n}...")
                            elif ev_tipo == "node_finished":
                                nd = ev.get("data", {})
                                if nd.get("status") == "failed":
                                    erro_fluxo = f"Falha no nó {nd.get('title')}: {nd.get('error')}"
                        except Exception:
                            pass

                    if erro_fluxo:
                        status.update(label="Falha no processamento", state="error")
                        st.error(f"Erro no Dify: {erro_fluxo}")
                    elif not outputs_finais:
                        status.update(label="Processamento sem saída", state="error")
                        st.error("O fluxo concluiu sem gerar os dados de saída.")
                    else:
                        res_perfis = outputs_finais.get("resultado_perfis", [])
                        if not res_perfis:
                            status.update(label="Sem dados gerados", state="error")
                            st.warning("A variável resultado_perfis retornou vazia.")
                        else:
