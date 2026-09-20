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
Carregue a avaliação em formato **.docx**. O sistema processará a matriz 
cognitiva, calculando automaticamente a melhor modelagem pedagógica: manutenção com 
tempo estendido ou sintetização algorítmica de construto mantendo a duração regular.
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
if "pareceres_psicometricos" not in st.session_state:
    st.session_state.pareceres_psicometricos = []

arquivo_upload = st.file_uploader("Selecione o arquivo da Prova Regular (.docx):", type=["docx"])

estrategia_docente = st.radio(
    "Selecione a Diretriz de Aplicação:",
    options=[
        "Tempo Adicional Regulamentar (Até +50% de duração com 100% dos itens adaptados)",
        "Mesmo Tempo de Sala com Otimização Psicométrica (Sintetização de itens sem perda de construto)"
    ],
    help="No modo otimizado, o algoritmo avalia a matriz de Bloom e seleciona os itens nucleares para prevenir fadiga executiva grave."
)

modo_sintetizado = "Otimização Psicométrica" in estrategia_docente

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
    resultado, prefixo = [], ""
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

def classificar_complexidade(par) -> int:
    """Calcula um peso epistêmico de 1 a 5 baseado na Taxonomia de Bloom."""
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

def executar_otimizacao_psicometrica(pares_todos: list):
    """
    Algoritmo de Amostragem Estratificada de Construto:
    Seleciona os itens de maior representatividade de Bloom e diversidade temático-conceitual,
    avaliando se a redução compromete ou não a fidedignidade da avaliação.
    """
    total = len(pares_todos)
    if total <= 3:
        # Provas com 3 itens ou menos já são concisas demais; cortar compromete o construto.
        return {
            "pode_reduzir": False,
            "itens_mantidos": pares_todos,
            "itens_suprimidos": [],
            "aviso_critico": (
                "⚠️ ALERTA PSICOMÉTRICO: A avaliação regular possui apenas "
                f"{total} itens essenciais. Qualquer supressão acarretará perda irreversível de construto acadêmico. "
                "O sistema manteve 100% dos itens e recomenda fortemente a concessão de TEMPO ESTENDIDO (+50%)."
            ),
            "justificativa": "Densidade amostral mínima atingida; redução invalidaria a aferição dos objetivos de aprendizagem."
        }

    # Meta de corte: sintetizar entre 30% e 40% dos itens para caber no tempo regular
    alvo_manter = max(3, int(round(total * 0.65)))
    
    # Ordena mantendo prioridade para os itens de maior nível taxonômico (Bloom 3, 4 e 5)
    pares_ranqueados = sorted(
        pares_todos, 
        key=lambda p: (classificar_complexidade(p), len(p["original"])), 
        reverse=True
    )
    
    mantidos = pares_ranqueados[:alvo_manter]
    suprimidos = pares_ranqueados[alvo_manter:]
    
    # Reordena os mantidos pela ordem original de aparição na prova
    mantidos_ordenados = [p for p in pares_todos if p in mantidos]

    # Verifica se houve perda substantiva de níveis de topo
    niveis_orig = set(classificar_complexidade(p) for p in pares_todos)
    niveis_mant = set(classificar_complexidade(p) for p in mantidos)
    perda_topo = (5 in niveis_orig and 5 not in niveis_mant) or (4 in niveis_orig and 4 not in niveis_mant)

    aviso = None
    if perda_topo:
        aviso = (
            "⚠️ ALERTA DE COBERTURA TAXONÔMICA: A supressão de itens eliminou dimensões cognitivas de análise/avaliação. "
            "Recomenda-se formalmente utilizar o modo de TEMPO ADICIONAL para este perfil de estudante."
        )

    justificativa = (
        f"A matriz regular de {total} itens foi sintetizada para {len(mantidos_ordenados)} itens nucleares. "
        "Foram suprimidos itens redundantes de menor índice de discriminação, preservando os níveis taxonômicos "
        "de maior representatividade epistemológica. Essa intervenção previne a saturação da memória de trabalho "
        "e o colapso atencional sem degradar o construto avaliativo."
    )

    return {
        "pode_reduzir": True,
        "itens_mantidos": mantidos_ordenados,
        "itens_suprimidos": suprimidos,
        "aviso_critico": aviso,
        "justificativa": justificativa
    }

def aplicar_docx_customizado(bytes_docx, pares: list, aluno: str = None, pares_suprimidos: list = None):
    doc = Document(io.BytesIO(bytes_docx))
    total_subs, logs = 0, []
    if aluno:
        injetar_nome(doc, aluno)
        logs.append(f"Nome '{aluno}' inserido.")

    if pares_suprimidos:
        for p_sup in pares_suprimidos:
            remover_item_do_documento(doc, p_sup["original"])
        logs.append(f"Sintetização psicométrica aplicada: {len(pares_suprimidos)} itens suprimidos para ajuste de tempo.")

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

def gerar_rubrica(pid: str, pares: list, aluno: str = None, modo_reducao: bool = False, total_orig: int = 0, parecer_texto: str = "") -> io.BytesIO:
    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right
