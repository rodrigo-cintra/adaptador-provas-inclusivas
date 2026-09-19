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
cognitiva e gerará os cadernos nominais, rubricas analíticas e protocolos 
de aplicação individualizados por estudante em um pacote único compactado.
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

arquivo_upload = st.file_uploader(
    "Selecione o arquivo da Prova Regular (.docx):", 
    type=["docx"]
)

contexto_turma_padrao = """[
  {"perfil_id": "TDAH_01", "alunos": ["Lucas Silva", "Gabriel Santos"]},
  {"perfil_id": "TEA_SUPORTE1", "alunos": ["Beatriz Mendes"]}
]"""

contexto_turma = st.text_area(
    "Mapeamento de Perfis da Turma (JSON):",
    value=contexto_turma_padrao,
    height=120
)

def sanitizar_nome_arquivo(nome: str) -> str:
    return re.sub(r'[^a-zA-Z0-9_-]', '_', str(nome).strip())

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
        preencher_paragrafo_com_markdown(
            paragrafo, 
            texto_p.replace(texto_antigo, texto_novo)
        )
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

def injetar_nome_no_cabecalho(doc: Document, nome_estudante: str) -> bool:
    """Substitui campos de identificação (Nome, Aluno, Discente) ou insere banner nominal."""
    padroes_busca = [
        r'(Nome\s*(?:do\s*Aluno\(a\)|do\s*Estudante|do\s*Aluno|Completo)?\s*:\s*)(_{2,}|\.{2,}|\s*)',
        r'(Aluno\(a\)\s*:\s*)(_{2,}|\.{2,}|\s*)',
        r'(Discente\s*:\s*)(_{2,}|\.{2,}|\s*)',
        r'(Estudante\s*:\s*)(_{2,}|\.{2,}|\s*)'
    ]

    def tentar_substituicao(p):
        txt = p.text
        for padrao in padroes_busca:
            if re.search(padrao, txt, re.IGNORECASE):
                # Substitui a linha preservando o rótulo e inserindo o nome em destaque
                def repl(m):
                    rotulo = m.group(1).strip()
                    return f"{rotulo} {nome_estudante}"
                novo_txt = re.sub(padrao, repl, txt, count=1, flags=re.IGNORECASE)
                p.text = ""
                r = p.add_run(novo_txt)
                r.bold = True
                r.font.color.rgb = RGBColor(24, 43, 73)
                return True
        return False

    # 1. Varre parágrafos das tabelas (comum em cabeçalhos institucionais)
    for tabela in doc.tables:
        for row in tabela.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if tentar_substituicao(p):
                        return True

    # 2. Varre os primeiros 20 parágrafos do corpo do documento
    for p in doc.paragraphs[:20]:
        if tentar_substituicao(p):
            return True

    # 3. Varre headers oficiais de seção
    for s in doc.sections:
        for p in s.header.paragraphs:
            if tentar_substituicao(p):
                return True

    # 4. Se não houver campo explícito, adiciona um banner elegante no início da página
    if doc.paragraphs:
        p_banner = doc.paragraphs[0].insert_paragraph_before()
    else:
        p_banner = doc.add_paragraph()
    
    r_rotulo = p_banner.add_run("Estudante: ")
    r_rotulo.bold = True
    r_rotulo.font.size = Pt(11)
    r_aluno = p_banner.add_run(f"{nome_estudante}\n")
    r_aluno.bold = True
    r_aluno.font.size = Pt(12)
    r_aluno.font.color.rgb = RGBColor(24, 43, 73)
    return True

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
        orig = (
            elem.get("texto_original") or 
            elem.get("enunciado_original") or 
            elem.get("original") or ""
        )
        adapt = (
            elem.get("texto_adaptado") or 
            elem.get("enunciado_adaptado") or 
            elem.get("adaptado") or ""
        )
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

def aplicar_adaptacoes_docx(bytes_docx_original, lista_pares: list, nome_aluno: str = None):
    doc = Document(io.BytesIO(bytes_docx_original))
    total_substituicoes = 0
    relatorio = []

    if nome_aluno:
        injetar_nome_no_cabecalho(doc, nome_aluno)
        relatorio.append(f"Nome do estudante '{nome_aluno}' inserido no cabeçalho.")

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

    barreira = "Sobrecarga de memória operacional decorrente de enunciado denso."
    ajuste = "Segmentação em comandos unitários com destaque visual nos verbos de ação."
    criterio_perfil = "Aceitar respostas sintéticas em tópicos, priorizando o rigor do conceito."

    if "TEA" in str(perfil_id).upper():
        barreira = "Ambiguidade na interpretação de comandos múltiplos e termos implícitos."
        ajuste = "Linearização dos comandos, vocabulário direto e eliminação de duplos sentidos."
        criterio_perfil = "Valorizar respostas literais e diretas, sem exigir floreios discursivos."

    return {
        "bloom": nivel_bloom,
        "barreira": dados.get("barreira_enfrentada") or barreira,
        "ajuste": dados.get("justificativa_acessibilidade") or ajuste,
        "criterio_perfil": dados.get("criterio_especifico") or criterio_perfil
    }

def gerar_rubrica_analitica_sofisticada(perfil_id: str, lista_pares: list, nome_aluno: str = None) -> io.BytesIO:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    p_title = doc.add_paragraph()
    run_title = p_title.add_run("Gabarito Orientado & Matriz de Correção Analítica")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(18)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(24, 43, 73)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub_texto = f"Estudante: {nome_aluno} | Perfil Funcional: {perfil_id}" if nome_aluno else f"Perfil Funcional: {perfil_id}"
    p_sub = doc.add_paragraph()
    run_sub = p_sub.add_run(sub_texto)
    run_sub.font.name = 'Calibri'
    run_sub.font.size = Pt(11)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(80, 80, 80)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    doc.add_heading("1. Fundamentação Pedagógica & Princípios Avaliativos", level=1)
    p_fund = doc.add_paragraph(
        "Este documento estabelece a matriz de correção técnica para o caderno adaptado, assegurando o princípio "
        "da equivalência cognitiva preconizado pelo Desenho Universal para a Aprendizagem (DUA) e pelas diretrizes "
        "institucionais de acessibilidade acadêmica. A adaptação visa eliminar barreiras instrumentais de acesso "
        "ao enunciado, mantendo integral o construto avaliativo e o rigor conceitual esperado para o componente curricular."
    )
    p_fund.paragraph_format.line_spacing = 1.15
    p_fund.paragraph_format.space_after = Pt(12)

    doc.add_heading("2. Matriz Analítica de Correção por Item", level=1)

    for idx, par in enumerate(lista_pares):
        num = par.get("numero", idx + 1)
        meta = inferir_metadados_psicometricos(par, perfil_id)

        h2 = doc.add_heading(f"Item #{num} — Análise Cognitiva & Critérios", level=2)
        h2.paragraph_format.space_before = Pt(14)
        h2.paragraph_format.space_after = Pt(4)

        tabela = doc.add_table(rows=6, cols=2)
        tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
        tabela.autofit = False

        largura_rotulo = Inches(2.0)
        largura_conteudo = Inches(4.5)

        dados_tabela = [
            ("Nível Taxonômico (Bloom):", meta["bloom"]),
            ("Enunciado Original:", par["original"]),
            ("Enunciado Adaptado:", par["adaptado"]),
            ("Barreira Mitigada:", meta["barreira"]),
            ("Intervenção Aplicada:", meta["ajuste"]),
            ("Diretriz Docente:", meta["criterio_perfil"])
        ]

        for i, (rotulo, valor) in enumerate(dados_tabela):
            linha = tabela.rows[i]
            c_rotulo = linha.cells[0]
            c_rotulo.width = largura_rotulo
            p_r = c_rotulo.paragraphs[0]
            r_run = p_r.add_run(rotulo)
            r_run.font.name = 'Calibri'
            r_run.font.size = Pt(9.5)
            r_run.font.bold = True
            definir_celula_fundo(c_rotulo, "F0F2F5")

            c_val = linha.cells[1]
            c_val.width = largura_conteudo
            p_v = c_val.paragraphs[0]
            v_run = p_v.add_run(str(valor))
            v_run.font.name = 'Calibri'
            v_run.font.size = Pt(9.5)

        doc.add_paragraph().paragraph_format.space_after = Pt(4)

        doc.add_heading(f"Rubrica de Desempenho Gradual — Item #{num}", level=3)
        tab_rubrica = doc.add_table(rows=4, cols=3)
        tab_rubrica.alignment = WD_TABLE_ALIGNMENT.CENTER
        tab_rubrica.autofit = False

        headers = ["Nível", "Critérios Conceituais Observáveis", "Ponderação"]
        larguras_rubrica = [Inches(1.8), Inches(3.6), Inches(1.1)]

        for c_idx, texto_h in enumerate(headers):
            cel = tab_rubrica.rows[0].cells[c_idx]
            cel.width = larguras_rubrica[c_idx]
            p_h = cel.paragraphs[0]
            r_h = p_h.add_run(texto_h)
            r_h.font.name = 'Calibri'
            r_h.font.size = Pt(9)
            r_h.font.bold = True
            r_h.font.color.rgb = RGBColor(255, 255, 255)
            definir_celula_fundo(cel, "1F3864")

        niveis_data = [
            (
                "Desempenho Pleno",
                "Mobiliza com precisão os conceitos solicitados nos comandos segmentados. Identifica e explica os elementos requeridos com clareza objetiva.",
                "90% a 100%"
            ),
            (
                "Desempenho Parcial",
                "Demonstra compreensão do núcleo central, respondendo corretamente à maioria dos subcomandos com omissão pontual de elementos secundários.",
                "50% a 70%"
            ),
            (
                "Desempenho Insuficiente",
                "Apresenta equívocos conceituais substantivos, fuga ao tema ou ausência de articulação entre os conceitos requeridos.",
                "0% a 30%"
            )
        ]

        for r_idx, (nivel, desc, pond) in enumerate(niveis_data, start=1):
            row = tab_rubrica.rows[r_idx]
            valores = [nivel, desc, pond]
            for c_idx, val in enumerate(valores):
                cel = row.cells[c_idx]
                cel.width = larguras_rubrica[c_idx]
                p = cel.paragraphs[0]
                run = p.add_run(val)
                run.font.name = 'Calibri'
                run.font.size = Pt(8.5)
                if c_idx == 0:
                    run.font.bold = True
                definir_celula_fundo(cel, "FFFFFF" if r_idx % 2 != 0 else "F9FAFC")

        doc.add_paragraph().paragraph_format.space_after = Pt(12)

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

def gerar_protocolo_aplicacao_avancado(perfil_id: str, lista_pares: list, nome_aluno: str = None) -> io.BytesIO:
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    p_t = doc.add_paragraph()
    r_t = p_t.add_run("Protocolo Oficial de Aplicação & Mediação Avaliativa")
    r_t.font.name = 'Calibri'
    r_t.font.size = Pt(18)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(24, 43, 73)
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub_txt = f"Estudante: {nome_aluno} | Perfil Funcional: {perfil_id}" if nome_aluno else f"Diretrizes de Sala | Perfil: {perfil_id}"
    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run(sub_txt)
    r_sub.font.name = 'Calibri'
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(80, 80, 80)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    doc.add_heading("1. Ficha de Parametrização & Registro de Sala", level=1)
    tab_ficha = doc.add_table(rows=5, cols=2)
    tab_ficha.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab_ficha.autofit = False

    larguras_ficha = [Inches(2.2), Inches(4.3)]
    dados_ficha = [
        ("Estudante Beneficiário:", nome_aluno if nome_aluno else "Conforme lista homologada"),
        ("Perfil Funcional Alvo:", f"{perfil_id} (Adaptação com Equivalência Cognitiva DUA)"),
        ("Fundamentação de Acessibilidade:", "Diretrizes Institucionais de Equidade e Apoio Didático Especializado"),
        ("Responsável pela Aplicação / Fiscal:", "________________________________________________________"),
        ("Horário Previsto (Início / Término):", "[ ___ : ___ ]  às  [ ___ : ___ ] (com acréscimo temporal)")
    ]

    for i, (campo, val) in enumerate(dados_ficha):
        row = tab_ficha.rows[i]
        c0 = row.cells[0]
        c0.width = larguras_ficha[0]
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(campo)
        r0.font.name = 'Calibri'
        r0.font.size = Pt(9.5)
        r0.font.bold = True
        definir_celula_fundo(c0, "F0F2F5")

        c1 = row.cells[1]
        c1.width = larguras_ficha[1]
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(val)
        r1.font.name = 'Calibri'
        r1.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    doc.add_heading("2. Matriz Operacional de Acomodações Avaliativas", level=1)
    tab_acomod = doc.add_table(rows=4, cols=3)
    tab_acomod.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab_acomod.autofit = False

    headers_ac = ["Dimensão", "Parâmetro Homologado", "Justificativa Neurocognitiva"]
    larguras_ac = [Inches(1.8), Inches(2.6), Inches(2.1)]

    for c_idx, h_txt in enumerate(headers_ac):
        cel = tab_acomod.rows[0].cells[c_idx]
        cel.width = larguras_ac[c_idx]
        p = cel.paragraphs[0]
        run = p.add_run(h_txt)
        run.font.name = 'Calibri'
        run.font.size = Pt(9.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        definir_celula_fundo(cel, "1F3864")

    tempo_just = "Compensa o tempo de processamento grafo-motor sem alterar o rigor acadêmico."
    espaco_param = "Assento nas primeiras fileiras ou sala com menor densidade de estímulos concorrentes."
    espaco_just = "Previne sobrecarga sensorial e descontinuidade atencional frente a ruídos externos."
    
    if "TDAH" in str(perfil_id).upper():
        pausa_param = "Pausas breves de autorregulação (3 a 5 min) a cada 45 min de execução."
        pausa_just = "Favorece o restabelecimento da memória de trabalho e mitiga o desgaste executivo."
    elif "TEA" in str(perfil_id).upper():
        pausa_param = "Pausas programadas para descompressão sensorial em ambiente calmo."
        pausa_just = "Previne crises de sobrecarga sensorial ou exaustão por saturação do ambiente."
    else:
        pausa_param = "Pausas regulares a critério do estudante com registro em ata."
        pausa_just = "Favorece a manutenção do foco e a precisão na escrita."

    acomod_linhas = [
        ("Acréscimo Temporal", "Até 50% de tempo adicional além da duração regular.", tempo_just),
        ("Ergonomia Espacial", espaco_param, espaco_just),
        ("Pausas Programadas", pausa_param, pausa_just)
    ]

    for r_idx, (dim, parm, just) in enumerate(acomod_linhas, start=1):
        row = tab_acomod.rows[r_idx]
        for c_idx, txt in enumerate([dim, parm, just]):
            cel = row.cells[c_idx]
            cel.width = larguras_ac[c_idx]
            p = cel.paragraphs[0]
            run = p.add_run(txt)
            run.font.name = 'Calibri'
            run.font.size = Pt(8.5)
            if c_idx == 0:
                run.font.bold = True
            definir_celula_fundo(cel, "FFFFFF" if r_idx % 2 != 0 else "F9FAFC")

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    doc.add_heading("3. Limiares de Mediação: Condutas Autorizadas e Vedações", level=1)
    tab_mediacao = doc.add_table(rows=5, cols=2)
    tab_mediacao.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab_mediacao.autofit = False

    larguras_med = [Inches(3.25), Inches(3.25)]
    h_med = ["Condutas Autorizadas (Mediação Válida)", "Condutas Vedadas (Compromete Fidedignidade)"]
    for c_idx, h_txt in enumerate(h_med):
        cel = tab_mediacao.rows[0].cells[c_idx]
        cel.width = larguras_med[c_idx]
        p = cel.paragraphs[0]
        run = p.add_run(h_txt)
        run.font.name = 'Calibri'
        run.font.size = Pt(9.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        definir_celula_fundo(cel, "2E75B6" if c_idx == 0 else "C00000")

    regras_mediacao = [
        (
            "Leitura Mediada de Enunciados: Reler comandos pausadamente, atendo-se ao texto do caderno adaptado.",
            "Parafrasear Conceitos: É vedado reescrever ou dar pistas sobre o tema acadêmico cobrado."
        ),
        (
            "Esclarecer Termos de Comando: Permitido clarificar o verbo de ação metodológico (ex.: 'relacione').",
            "Validar Respostas: Estritamente proibido dizer 'está no caminho certo' ou sugerir revisões induzidas."
        ),
        (
            "Recursos de Apoio: Permitir folhas de rascunho sem limites, marca-textos coloridos e réguas guia.",
            "Impor Padrão Gráfico Rígido: Não coagir o estudante a escrever em prosa contínua se ele optar por tópicos."
        ),
        (
            "Sinalização Individual de Tempo: Avisar o tempo restante de forma discreta, sem alarmar a turma toda.",
            "Pressão por Celeridade: Nunca induzir o aluno a acelerar com comparações ao término de outros colegas."
        )
    ]

    for r_idx, (permitido, proibido) in enumerate(regras_mediacao, start=1):
        row = tab_mediacao.rows[r_idx]
        for c_idx, val in enumerate([permitido, proibido]):
            cel = row.cells[c_idx]
            cel.width = larguras_med[c_idx]
            p = cel.paragraphs[0]
            run = p.add_run(val)
            run.font.name = 'Calibri'
            run.font.size = Pt(8.5)
            definir_celula_fundo(cel, "F2F7FA" if c_idx == 0 else "FDF2F2")

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    doc.add_heading("4. Protocolo de Contingência: Sobrecarga & Ansiedade", level=1)
    p_crise = doc.add_paragraph(
        "Caso o estudante manifeste sinais de bloqueio atencional, agitação psicomotora, sobrecarga sensorial ou crise ansiosa:\n"
        "1. Interrupção Neutra: Convide o estudante a afastar-se momentaneamente da folha de prova de forma discreta.\n"
        "2. Pausa Hídrica / Descompressão: Permita ingestão de água e respiração consciente fora da sala por até 5 minutos.\n"
        "3. Retomada sem Pressão: Ao regressar, oriente-o a reiniciar por um item considerado mais acessível.\n"
        "4. Registro Obrigatório: Qualquer intercorrência atípica deve ser descrita no Termo de Encerramento abaixo."
    )
    p_crise.paragraph_format.line_spacing = 1.15
    p_crise.paragraph_format.space_after = Pt(10)

    doc.add_heading("5. Termo de Encerramento & Conformidade", level=1)
    tab_termo = doc.add_table(rows=3, cols=1)
    tab_termo.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab_termo.autofit = False
    
    cel_obs = tab_termo.rows[0].cells[0]
    cel_obs.width = Inches(6.5)
    p_obs = cel_obs.paragraphs[0]
    p_obs.add_run("Registro de Ocorrências / Observações Relevantes de Sala:\n\n\n").font.size = Pt(9)
    definir_celula_fundo(cel_obs, "FAFAFA")

    cel_decl = tab_termo.rows[1].cells[0]
    cel_decl.width = Inches(6.5)
    p_decl = cel_decl.paragraphs[0]
    p_decl.add_run(
        "Declaro que a avaliação foi administrada em estrita conformidade com os parâmetros operacionais e éticos "
        "deste protocolo, assegurando a lisura acadêmica e os direitos de acessibilidade do estudante."
    ).font.size = Pt(8.5)

    cel_ass = tab_termo.rows[2].cells[0]
    cel_ass.width = Inches(6.5)
    p_ass = cel_ass.paragraphs[0]
    p_ass.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ass.add_run(
        "\n___________________________________________________\n"
        "Assinatura do Responsável pela Aplicação / Fiscal de Sala\n"
        "Data: _____ / _____ / 2026"
    ).font.size = Pt(9)

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

                # Mapeamento prévio dos perfis configurados no JSON de entrada
                lista_perfis_config = []
                try:
                    turma_parsed = json.loads(contexto_turma)
                    if isinstance(turma_parsed, list):
                        lista_perfis_config = turma_parsed
                except Exception:
                    lista_perfis_config = []

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
                    st.write("Adaptando questões e gerando cadernos nominais por estudante...")

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
                                    st.write(f"Etapa: {no_nome}...")
                            elif evento == "node_finished":
                                dados_no = dados_evento.get("data", {})
                                if dados_no.get("status") == "failed":
                                    erro_fluxo = f"Falha no nó {dados_no.get('title')}: {dados_no.get('error')}"
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
                            st.warning("A variável resultado_perfis retornou vazia.")
                        else:
                            buffer_zip = io.BytesIO()

                            with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zip_file:
                                total_cadernos_gerados = 0

                                for idx, item_perfil in enumerate(resultado_perfis):
                                    # Associação resiliente: tenta por índice posicional e por perfil_id
                                    perfil_config_atual = {}
                                    if idx < len(lista_perfis_config):
                                        perfil_config_atual = lista_perfis_config[idx]

                                    p_id = perfil_config_atual.get("perfil_id")
                                    if not p_id and isinstance(item_perfil, dict):
                                        p_id = item_perfil.get("perfil_id")
                                    if not p_id:
                                        p_id = f"PERFIL_{idx + 1}"

                                    # Obtém os alunos configurados
                                    alunos_perfil = perfil_config_atual.get("alunos", [])
                                    if not alunos_perfil and isinstance(item_perfil, dict):
                                        alunos_perfil = item_perfil.get("alunos", [])
                                    if not alunos_perfil:
                                        alunos_perfil = [f"Estudante_{p_id}"]

                                    pares_adaptacao = extrair_dados_perfil(item_perfil)
                                    total_subs_perfil = 0

                                    # Cria uma pasta e um kit completo por estudante cadastrado
                                    for aluno in alunos_perfil:
                                        pasta_estudante = sanitizar_nome_arquivo(f"{aluno}_{p_id}")
                                        total_cadernos_gerados += 1

                                        # 1. Caderno de Prova Adaptada Nominal (.docx)
                                        docx_adaptado, total_subs, relatorio = aplicar_adaptacoes_docx(
                                            bytes_docx, 
                                            pares_adaptacao, 
                                            nome_aluno=aluno
                                        )
                                        total_subs_perfil = total_subs
                                        nome_prova = f"{pasta_estudante}/Caderno_Prova_{sanitizar_nome_arquivo(aluno)}.docx"
                                        zip_file.writestr(nome_prova, docx_adaptado.getvalue())

                                        # 2. Gabarito & Matriz Analítica Nominal (.docx)
                                        docx_gabarito = gerar_rubrica_analitica_sofisticada(
                                            p_id, 
                                            pares_adaptacao, 
                                            nome_aluno=aluno
                                        )
                                        nome_gabarito = f"{pasta_estudante}/Gabarito_e_Rubrica_{sanitizar_nome_arquivo(aluno)}.docx"
                                        zip_file.writestr(nome_gabarito, docx_gabarito.getvalue())

                                        # 3. Protocolo de Aplicação Nominal (.docx)
                                        docx_instrucoes = gerar_protocolo_aplicacao_avancado(
                                            p_id, 
                                            pares_adaptacao, 
                                            nome_aluno=aluno
                                        )
                                        nome_instrucoes = f"{pasta_estudante}/Protocolo_Aplicacao_{sanitizar_nome_arquivo(aluno)}.docx"
                                        zip_file.writestr(nome_instrucoes, docx_instrucoes.getvalue())

                                    st.session_state.resumo_geracao.append({
                                        "perfil": p_id,
                                        "estudantes": alunos_perfil,
                                        "alteracoes": total_subs_perfil,
                                        "total_pares": len(pares_adaptacao)
                                    })
                                    st.session_state.detalhes_log.append({
                                        "perfil": p_id,
                                        "estudantes": alunos_perfil,
                                        "pares": pares_adaptacao
                                    })

                            buffer_zip.seek(0)
                            st.session_state.pacote_zip = buffer_zip.getvalue()
                            status.update(
                                label=f"Sucesso! {total_cadernos_gerados} cadernos nominais gerados no pacote.", 
                                state="complete"
                            )

            except Exception as e:
                status.update(label="Erro no processamento", state="error")
                st.error(f"Ocorreu um erro: {str(e)}")

if st.session_state.pacote_zip:
    st.divider()
    st.subheader("📦 Pacote Pedagógico Pronto para Download")
    st.markdown("O arquivo compactado organiza **uma pasta nominal para cada estudante** cadastrado:")
    st.markdown("- **Caderno de Prova Adaptado e Nominal** (`.docx` com nome do estudante inserido no cabeçalho)")
    st.markdown("- **Gabarito & Matriz de Correção Nominal** (`.docx` com psicometria e rubrica por aluno)")
    st.markdown("- **Protocolo Oficial de Aplicação Nominal** (`.docx` com ficha de sala e registro de acomodações)")

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
        cols_metrica[i].metric(
            label=f"Perfil: {r['perfil']} ({len(r['estudantes'])} alunos)",
            value=f"{r['alteracoes']} modificadas",
            help=f"Estudantes atendidos: {alunos_str}"
        )

    with st.expander("🔍 Auditoria detalhada dos estudantes e substituições"):
        for d in st.session_state.detalhes_log:
            st.markdown(f"### Perfil: {d['perfil']}")
            st.markdown(f"**Estudantes Gerados:** {', '.join(d['estudantes'])}")
            st.markdown("**Pares aplicados nas questões:**")
            st.json(d['pares'])
