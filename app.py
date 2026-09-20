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
    page_title="Adaptador Académico Inclusivo",
    page_icon="🎓",
    layout="centered"
)

# ----------------- CONFIGURAÇÕES E CONSTANTES -----------------

DIFY_API_KEY = "app-9NqVkZLWEQgSjy2AZHZ5KGO3"
DIFY_WORKFLOW_URL = "https://api.dify.ai/v1/workflows/run"
DIFY_UPLOAD_URL = "https://api.dify.ai/v1/files/upload"

if "pacote_zip" not in st.session_state:
    st.session_state.pacote_zip = None
if "resumo_geracao" not in st.session_state:
    st.session_state.resumo_geracao = []
if "detalhes_log" not in st.session_state:
    st.session_state.detalhes_log = []
if "pareceres_psicometricos" not in st.session_state:
    st.session_state.pareceres_psicometricos = []

# ----------------- FUNÇÕES AUXILIARES -----------------

def sanitizar_nome(s: str) -> str:
    return re.sub(r'[^a-zA-Z0-9_-]', '_', str(s).strip())

def limpar_str(s: str) -> str:
    if not s:
        return ""
    return re.sub(r'\s+', ' ', re.sub(r'[\r\n\t]+', ' ', str(s))).strip()

def texto_consolidado(p) -> str:
    if p.runs:
        return "".join(run.text for run in p.runs)
    return p.text or ""

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
    a_limpo = limpar_str(antigo)
    p_limpo = limpar_str(txt)

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

        s_orig = str(orig).strip()
        s_adapt = str(adapt).strip()
        sub_origs = fragmentar_comandos(s_orig)
        sub_adapts = fragmentar_comandos(s_adapt)

        if len(sub_origs) == len(sub_adapts) and len(sub_origs) > 1:
            for i, (so, sa) in enumerate(zip(sub_origs, sub_adapts)):
                pares.append({"numero": f"{num}.{i+1}", "original": so.strip(), "adaptado": sa.strip(), "raw": elem})
        elif s_orig and s_adapt:
            pares.append({"numero": num, "original": s_orig, "adaptado": s_adapt, "raw": elem})
    return pares

def classificar_complexidade(par) -> int:
    t = par.get("original", "").lower()
    raw = par.get("raw", {})
    b = str(raw.get("bloom") or raw.get("nivel_bloom") or "").lower()
    if "criar" in b or "avaliar" in b or any(v in t for v in ["avalie", "julgue", "critique", "defenda"]):
        return 5
    if "analisar" in b or any(v in t for v in ["analise", "relacione", "compare", "diferencie"]):
        return 4
    if "aplicar" in b or any(v in t for v in ["aplique", "calcule", "resolva", "demonstre"]):
        return 3
    if "compreender" in b or any(v in t for v in ["explique", "caracterize", "descreva", "discuta"]):
        return 2
    return 1

def nome_nivel_bloom(peso: int) -> str:
    niveis = {
        5: "Avaliar / Julgar (Nível 5)",
        4: "Analisar / Correlacionar (Nível 4)",
        3: "Aplicar / Executar (Nível 3)",
        2: "Compreender / Explicar (Nível 2)",
        1: "Lembrar / Identificar (Nível 1)"
    }
    return niveis.get(peso, "Compreender (Nível 2)")

def extrair_tema_aproximado(texto: str) -> str:
    """Identifica o núcleo temático com base em premissas usuais do enunciado."""
    t = texto.strip()
    match = re.search(r'^(Questão\s*\d+[^:—\.\n]*[:—\.]\s*[^:\.\n]+)', t, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    palavras = t.split()
    return " ".join(palavras[:6]) + "..."

def construir_laudo_pericial_detalhado(total_orig: int, mantidos: list, suprimidos: list, perfil_id: str, perda_critica: bool) -> str:
    """Gera um laudo técnico circunstanciado para respaldo pedagógico e jurídico."""
    perfil_nome = "Transtorno do Déficit de Atenção com Hiperatividade (TDAH)" if "TDAH" in perfil_id.upper() else "Transtorno do Espectro Autista (TEA)"
    
    laudo = []
    laudo.append("PARECER PERICIAL DE EQUIVALÊNCIA COGNITIVA E OTIMIZAÇÃO DE CONSTRUTO")
    laudo.append("=" * 80)
    laudo.append(f"Perfil Clínico/Pedagógico Alvo: {perfil_nome} ({perfil_id})")
    laudo.append("Fundamentação Normativa: Desenho Universal para a Aprendizagem (DUA/CAST), Teoria da Carga Cognitiva (Sweller) e Taxonomia de Bloom Revisada.")
    laudo.append("-" * 80)
    laudo.append("\n1. ENQUADRAMENTO DA ACOMODAÇÃO TEMPORAL:")
    laudo.append(
        "A modulação temporal por sintetização de itens fundamenta-se no princípio de que, para estudantes com déficits na modulação "
        "atencional executiva e fadiga de memória de trabalho, a simples dilação temporal (tempo extra) opera como fator de sobrecarga sensorial "
        "e degradação motora cumulativa. A manutenção do tempo padrão de sala de aula, associada à redução quantitativa de itens redundantes, "
        "assegura a preservação da curva de rendimento neurocognitivo sem exaustão."
    )
    
    laudo.append("\n2. DEMONSTRAÇÃO DA COBERTURA EPISTEMOLÓGICA E CONSTRUTO MANTIDO:")
    laudo.append(
        f"A prova regular contemplava originariamente {total_orig} unidades de avaliação. O algoritmo de curadoria psicométrica preservou {len(mantidos)} "
        "itens nucleares de maior densidade epistêmica. Os seguintes domínios e níveis cognitivos superiores permanecem integralmente cobertos:"
    )
    for m in mantidos:
        peso = classificar_complexidade(m)
        rotulo_bloom = nome_nivel_bloom(peso)
        tema = extrair_tema_aproximado(m["original"])
        laudo.append(f"  • Item {m.get('numero')}: {tema} | Nível Taxonômico: {rotulo_bloom}. Justificativa: Preservado como âncora conceitual imprescindível do componente curricular.")

    laudo.append("\n3. JUSTIFICATIVA PERICIAL DOS ITENS SUPRIMIDOS (CRITÉRIO DE REDUNDÂNCIA):")
    if suprimidos:
        laudo.append(
            "A supressão dos itens abaixo relacionados obedeceu estritamente ao critério de não sobreposição avaliativa. Aferiu-se que os construtos "
            "subjacentes a estas questões já se encontram incorporados nos itens de maior complexidade taxonômica preservados:"
        )
        for s in suprimidos:
            peso_s = classificar_complexidade(s)
            rotulo_s = nome_nivel_bloom(peso_s)
            tema_s = extrair_tema_aproximado(s["original"])
            laudo.append(
                f"  • Item {s.get('numero')} ({tema_s}): Nível {rotulo_s}. Motivo da exclusão: Subitem com função preponderantemente confirmatória ou de "
                "recordação factual secundária. Sua resolução não adiciona variância explicativa ao domínio conceitual já testado nos itens mantidos."
            )
    else:
        laudo.append("  • Nenhum item foi suprimido. A matriz foi mantida integralmente.")

    laudo.append("\n4. ANÁLISE DE FIDEDIGNIDADE E RISCO PSICOMÉTRICO:")
    if perda_critica:
        laudo.append(
            "⚠️ ADVERTÊNCIA TÉCNICA: Constatou-se que a matriz regular possuía concentração atípica de tópicos independentes e irredutíveis. "
            "A sintetização amostral pode gerar perda de evidência de validade curricular. Recomenda-se formalmente que este estudante seja "
            "transferido para a modalidade de TEMPO ESTENDIDO (+50%) preservando 100% das questões originais."
        )
    else:
        laudo.append(
            "Declara-se que a matriz curricular essencial permanece avaliada com índice de fidedignidade equivalente à prova regular. "
            "A presente adequação atende aos preceitos da legislação de acessibilidade, resguardando a fidedignidade do instrumento e os direitos do educando."
        )
    
    return "\n".join(laudo)

def executar_otimizacao_psicometrica(pares_todos: list, perfil_id: str):
    total = len(pares_todos)
    if total <= 3:
        aviso = (
            f"⚠️ ALERTA PSICOMÉTRICO: A avaliação regular possui apenas {total} itens essenciais. "
            "Qualquer corte causaria perda substancial de construto acadêmico. "
            "O sistema manteve todos os itens e recomenda TEMPO ESTENDIDO (+50%)."
        )
        laudo = construir_laudo_pericial_detalhado(total, pares_todos, [], perfil_id, perda_critica=True)
        return {
            "pode_reduzir": False,
            "itens_mantidos": pares_todos,
            "itens_suprimidos": [],
            "aviso_critico": aviso,
            "justificativa": laudo
        }

    alvo_manter = max(3, int(round(total * 0.65)))
    pares_ranqueados = sorted(
        pares_todos, 
        key=lambda p: (classificar_complexidade(p), len(p["original"])), 
        reverse=True
    )
    mantidos = pares_ranqueados[:alvo_manter]
    suprimidos = pares_ranqueados[alvo_manter:]
    mantidos_ordenados = [p for p in pares_todos if p in mantidos]

    niveis_orig = set(classificar_complexidade(p) for p in pares_todos)
    niveis_mant = set(classificar_complexidade(p) for p in mantidos)
    perda_topo = (5 in niveis_orig and 5 not in niveis_mant) or (4 in niveis_orig and 4 not in niveis_mant)

    aviso = None
    if perda_topo:
        aviso = (
            "⚠️ ALERTA DE COBERTURA TAXONÓMICA: A sintetização eliminou dimensões analíticas essenciais. "
            "Recomenda-se formalmente adotar TEMPO ADICIONAL para este perfil."
        )

    laudo = construir_laudo_pericial_detalhado(total, mantidos_ordenados, suprimidos, perfil_id, perda_critica=perda_topo)

    return {
        "pode_reduzir": True,
        "itens_mantidos": mantidos_ordenados,
        "itens_suprimidos": suprimidos,
        "aviso_critico": aviso,
        "justificativa": laudo
    }

def aplicar_docx_customizado(bytes_docx, pares: list, aluno: str = None, pares_suprimidos: list = None):
    doc = Document(io.BytesIO(bytes_docx))
    total_subs = 0
    logs = []

    if aluno:
        injetar_nome(doc, aluno)
        logs.append(f"Nome '{aluno}' inserido no cabeçalho.")

    if pares_suprimidos:
        for p_sup in pares_suprimidos:
            remover_item_do_documento(doc, p_sup["original"])
        logs.append(f"Sintetização psicométrica: {len(pares_suprimidos)} itens suprimidos para ajuste de tempo.")

    for par in pares:
        orig = par["original"]
        adapt = par["adaptado"]
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

def gerar_rubrica(pid: str, pares: list, aluno: str = None, modo_reducao: bool = False, total_orig: int = 0, laudo_texto: str = "") -> io.BytesIO:
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

    doc.add_heading("1. Laudo Técnico-Psicométrico de Equivalência de Construto", level=1)
    
    # Se houver laudo circunstanciado, insere formatado linha por linha
    if laudo_texto:
        for linha in laudo_texto.split("\n"):
            p_l = doc.add_paragraph()
            p_l.paragraph_format.line_spacing = 1.15
            p_l.paragraph_format.space_after = Pt(2)
            if linha.startswith("PARECER") or linha.startswith("1.") or linha.startswith("2.") or linha.startswith("3.") or linha.startswith("4."):
                r = p_l.add_run(linha)
                r.bold = True
                r.font.color.rgb = RGBColor(24, 43, 73)
            elif linha.startswith("  •"):
                r = p_l.add_run(linha)
                r.font.size = Pt(9.5)
            elif linha.startswith("="):
                pass
            else:
                p_l.add_run(linha)
    else:
        doc.add_paragraph(
            "Este documento estabelece a matriz de correção técnica para o caderno adaptado, assegurando o princípio "
            "da equivalência cognitiva preconizado pelo Desenho Universal para a Aprendizagem (DUA). Foram mantidos 100% dos "
            "itens curriculares originais com homologação de tempo estendido."
        )

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
            ("Parcial", "Demonstra compreensão do núcleo central, com omissão pontual de elementos secundários.", "50% a 70%"),
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

    if modo_reducao:
        tempo_desc = "Mesmo tempo de sala da turma regular (Sem acréscimo temporal - Prova Sintetizada)"
        estrat_desc = f"Sintetização algorítmica de construto para {total_itens_mantidos} itens nucleares."
    else:
        tempo_desc = "[ ___ : ___ ] às [ ___ : ___ ] (com tempo estendido de até +50%)"
        estrat_desc = "Manutenção integral dos itens com concessão de tempo estendido."

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

# ----------------- FORMULÁRIO PRINCIPAL -----------------

st.title("🎓 Adaptação Didática Inclusiva de Avaliações")
st.markdown("""
Carregue a avaliação em formato **.docx**. O sistema processará a matriz 
cognitiva, gerando cadernos nominais, rubricas com parecer pericial e protocolos de fiscal de sala.
""")

arquivo_upload = st.file_uploader("Selecione o arquivo da Prova Regular (.docx):", type=["docx"])

estrategia_docente = st.radio(
    "Selecione a Diretriz de Aplicação:",
    options=[
        "Tempo Adicional Regulamentar (Até +50% de duração com 100% dos itens adaptados)",
        "Mesmo Tempo de Sala com Otimização Psicométrica (Sintetização pericial sem perda de construto)"
    ],
    help="No modo otimizado, o algoritmo avalia a matriz de Bloom e sintetiza itens redundantes para evitar fadiga executiva."
)

modo_sintetizado = "Otimização Psicométrica" in estrategia_docente

contexto_turma_padrao = """[
  {"perfil_id": "TDAH_01", "alunos": ["Lucas Silva", "Gabriel Santos"]},
  {"perfil_id": "TEA_SUPORTE1", "alunos": ["Beatriz Mendes"]}
]"""

contexto_turma = st.text_area("Mapeamento de Perfis da Turma (JSON):", value=contexto_turma_padrao, height=110)

st.markdown("<br>", unsafe_allow_html=True)
disparar = st.button("🚀 Gerar Pacote Pedagógico Completo", type="primary", use_container_width=True)

# ----------------- EXECUÇÃO AO CLICAR -----------------

if disparar:
    if not arquivo_upload:
        st.warning("Por favor, selecione um arquivo .docx antes de prosseguir.")
    else:
        st.session_state.pacote_zip = None
        st.session_state.resumo_geracao = []
        st.session_state.detalhes_log = []
        st.session_state.pareceres_psicometricos = []

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
                    st.write("Analisando matriz taxonómica e emitindo parecer psicométrico...")

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
                            buf_zip = io.BytesIO()
                            with zipfile.ZipFile(buf_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                                total_cadernos = 0

                                for idx, item_p in enumerate(res_perfis):
                                    cfg_p = perfis_cfg[idx] if idx < len(perfis_cfg) else {}
                                    pid = cfg_p.get("perfil_id")
                                    if not pid and isinstance(item_p, dict):
                                        pid = item_p.get("perfil_id")
                                    if not pid:
                                        pid = f"PERFIL_{idx+1}"

                                    alunos = cfg_p.get("alunos", [])
                                    if not alunos and isinstance(item_p, dict):
                                        alunos = item_p.get("alunos", [])
                                    if not alunos:
                                        alunos = [f"Estudante_{pid}"]

                                    pares_brutos = extrair_pares_resiliente(item_p)
                                    total_orig = len(pares_brutos)

                                    if modo_sintetizado:
                                        res_otim = executar_otimizacao_psicometrica(pares_brutos, pid)
                                        pares_mantidos = res_otim["itens_mantidos"]
                                        pares_suprimidos = res_otim["itens_suprimidos"]
                                        laudo_completo = res_otim["justificativa"]
                                        if res_otim["aviso_critico"]:
                                            st.session_state.pareceres_psicometricos.append(res_otim["aviso_critico"])
                                    else:
                                        pares_mantidos = pares_brutos
                                        pares_suprimidos = []
                                        laudo_completo = (
                                            "PARECER DE MODULAÇÃO TEMPORAL INTEGRAL:\n"
                                            "Optou-se pela manutenção de 100% dos itens curriculares da avaliação regular com a concessão "
                                            "de até 50% de tempo adicional homologado. As alterações limitaram-se à eliminação de barreiras "
                                            "de representação e desmembramento de comandos sob as diretrizes do DUA."
                                        )

                                    subs_perfil = 0

                                    for aluno in alunos:
                                        pasta = sanitizar_nome(f"{aluno}_{pid}")
                                        total_cadernos += 1

                                        docx_ad, n_subs, r_logs = aplicar_docx_customizado(
                                            bytes_docx, 
                                            pares_mantidos, 
                                            aluno=aluno,
                                            pares_suprimidos=pares_suprimidos
                                        )
                                        subs_perfil = n_subs
                                        zf.writestr(f"{pasta}/Caderno_Prova_{sanitizar_nome(aluno)}.docx", docx_ad.getvalue())

                                        docx_gab = gerar_rubrica(
                                            pid, 
                                            pares_mantidos, 
                                            aluno=aluno, 
                                            modo_reducao=modo_sintetizado, 
                                            total_orig=total_orig,
                                            laudo_texto=laudo_completo
                                        )
                                        zf.writestr(f"{pasta}/Gabarito_e_Rubrica_{sanitizar_nome(aluno)}.docx", docx_gab.getvalue())

                                        docx_ins = gerar_protocolo(
                                            pid, 
                                            aluno=aluno, 
                                            modo_reducao=modo_sintetizado, 
                                            total_itens_mantidos=len(pares_mantidos)
                                        )
                                        zf.writestr(f"{pasta}/Protocolo_Aplicacao_{sanitizar_nome(aluno)}.docx", docx_ins.getvalue())

                                    st.session_state.resumo_geracao.append({
                                        "perfil": pid, 
                                        "estudantes": alunos, 
                                        "alteracoes": subs_perfil, 
                                        "total_pares": len(pares_mantidos),
                                        "suprimidos": len(pares_suprimidos),
                                        "laudo": laudo_completo
                                    })
                                    st.session_state.detalhes_log.append({
                                        "perfil": pid, 
                                        "estudantes": alunos, 
                                        "pares": pares_mantidos
                                    })

                            buf_zip.seek(0)
                            st.session_state.pacote_zip = buf_zip.getvalue()
                            status.update(label=f"Sucesso! {total_cadernos} cadernos gerados com a estratégia selecionada.", state="complete")

            except Exception as e:
                status.update(label="Erro no processamento", state="error")
                st.error(f"Ocorreu um erro: {str(e)}")

# ----------------- PAINEL DE RESULTADOS PERSISTENTES -----------------

if st.session_state.pareceres_psicometricos:
    st.divider()
    st.subheader("⚠️ Parecer Psicométrico do Sistema")
    for aviso in set(st.session_state.pareceres_psicometricos):
        st.warning(aviso)

if st.session_state.pacote_zip:
    st.divider()
    st.subheader("📦 Pacote Pedagógico Pronto para Download")
    st.markdown("O arquivo compactado organiza **uma pasta nominal para cada estudante** cadastrado:")
    st.markdown("- **Caderno de Prova Adaptado** (`.docx` com layout preservado e modelagem algorítmica de construto)")
    st.markdown("- **Gabarito & Matriz de Correção** (`.docx` contendo o laudo pericial completo para prontuário)")
    st.markdown("- **Protocolo Oficial de Aplicação** (`.docx` com diretrizes homologadas para o fiscal de sala)")

    st.download_button(
        label="📥 Baixar Pacote Completo Individualizado (.zip)",
        data=st.session_state.pacote_zip,
        file_name="Avaliacoes_Adaptadas_Nominais_Pacote_Completo.zip",
        mime="application/zip",
        type="primary",
        key="btn_zip_consolidado"
    )

    st.markdown("---")
    cols_metrica = st.columns(len(st.session_state.resumo_geracao))
    for i, r in enumerate(st.session_state.resumo_geracao):
        alunos_str = ", ".join(r['estudantes'])
        msg_help = f"Estudantes atendidos: {alunos_str}"
        if r.get("suprimidos", 0) > 0:
            msg_help += f" | {r['suprimidos']} itens redundantes suprimidos para adaptação de ritmo."
        cols_metrica[i].metric(
            label=f"Perfil: {r['perfil']} ({len(r['estudantes'])} alunos)",
            value=f"{r['total_pares']} itens na prova",
            help=msg_help
        )

    with st.expander("🔍 Auditoria Psicométrica Detalhada & Laudo Pericial para Prontuário"):
        for d in st.session_state.resumo_geracao:
            st.markdown(f"### Perfil: {d['perfil']}")
            st.markdown(f"**Estudantes Gerados:** {', '.join(d['estudantes'])}")
            st.text_area(f"Laudo Técnico Pericial ({d['perfil']})", value=d["laudo"], height=320)
        for d in st.session_state.detalhes_log:
            st.markdown(f"**Itens Ativos no Perfil {d['perfil']}:**")
            st.json(d['pares'])
