import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import fitz  # PyMuPDF
import io
import os
import subprocess
import tempfile
import pandas as pd
import matplotlib.pyplot as plt
from streamlit_paste_button import paste_image_button
from PIL import Image
import platform
import time
import calendar

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="Gerador de Relatórios Madalena", layout="wide")

# --- CUSTOM CSS ---
st.markdown("""
    <style>
    .main { background-color: #f0f2f5; }
    
    /* CONFIGURAÇÃO DO GHOST CARD VIA CONTAINER NATIVO */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #ffffff !important;
        border: 1px solid #d1d5db !important;
        border-radius: 12px !important;
        padding: 25px !important;
        margin-bottom: 20px !important;
    }
    
    div.stButton > button[kind="primary"] {
        background-color: #2c86b0 !important;
        color: white !important;
        border: none !important;
        width: 100% !important;
        font-weight: bold !important;
        height: 3em !important;
        border-radius: 8px !important;
    }
    div.stButton > button[key*="del_"] {
        border: 1px solid #dc3545 !important;
        color: #dc3545 !important;
        background-color: transparent !important;
        font-size: 0.8em !important;
        height: 2em !important;
    }
    .upload-label { font-weight: bold; color: #1f2937; margin-bottom: 8px; display: block; }
    </style>
    """, unsafe_allow_html=True)

# --- DICIONÁRIO DE DIMENSÕES DAS EVIDÊNCIAS ---
DIMENSOES_CAMPOS = {
    "H_GRAFICO_ATEND_AMB": 165, "H_GRAFICO_ASSIST_HOSP": 125, "H_TAB_TRANS_HOSP": 125,
    "H_GRAFICO_TRANS_HOSP": 125, "H_GRAFICO_SAIDA_HOSP": 125,
    "H_GRAFICO_ATEND_EMERG": 180, "ATA_REUNIAO_OBITO": 165, "ATA_REUNIAO_PRONTUARIO": 165,
    "ATA_REUNIAO_INFEC": 165, "CAPS_GRAFICO_ATEND": 165, "CAPS_REGISTRO_FOTOGRAFICO": 165,
    "AB_GRAFICO_ATEND": 175, "AB_METAQUANTI_HOSP": 130, "AB_METAQUALI_HOSP": 175
}

# --- ESTADO DA SESSÃO ---
if 'dados_sessao' not in st.session_state:
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3208/3208726.png", width=100)
    st.title("Painel de Controle")
    st.markdown("---")
    total_anexos = sum(len(v) for v in st.session_state.dados_sessao.values())
    st.metric("Total de Evidências", total_anexos)
    if st.button(" 🗑️ Limpar Todos os Dados", width='stretch'):
        st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
        st.rerun()

# --- FUNÇÕES CORE ---
def converter_para_pdf(docx_path, output_dir):
    comando = 'libreoffice'
    if platform.system() == "Windows":
        caminhos = ['libreoffice', r'C:\Program Files\LibreOffice\program\soffice.exe', r'C:\Program Files (x86)\LibreOffice\program\soffice.exe']
        for p in caminhos:
            try:
                subprocess.run([p, '--version'], capture_output=True, check=True)
                comando = p
                break
            except: continue
    subprocess.run([comando, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path], check=True)

def processar_item_lista(doc_template, item, marcador):
    largura = DIMENSOES_CAMPOS.get(marcador, 165)
    try:
        if isinstance(item, Image.Image):
            img_buf = io.BytesIO()
            item.save(img_buf, format='PNG')
            img_buf.seek(0)
            return [InlineImage(doc_template, img_buf, width=Mm(largura))]
        if hasattr(item, 'seek'): item.seek(0)
        ext = getattr(item, 'name', '').lower()
        if ext.endswith(".pdf"):
            pdf = fitz.open(stream=item.read(), filetype="pdf")
            imgs = []
            for pg in pdf:
                pix = pg.get_pixmap(matrix=fitz.Matrix(2, 2))
                imgs.append(InlineImage(doc_template, io.BytesIO(pix.tobytes()), width=Mm(largura)))
            pdf.close()
            return imgs
        return [InlineImage(doc_template, item, width=Mm(largura))]
    except Exception: return []

# --- UI PRINCIPAL ---
st.title("Automação de Relatórios - Madalena")
st.caption("Versão 0.9.2")

t_manual_amb, t_manual_caps, t_manual_ab, t_evidencia = st.tabs(["AMBULATORIAL", "CAPS", "ATENÇÃO BÁSICA", "ARQUIVOS"])

# --- ABA AMBULATORIAL ---
with t_manual_amb:
    with st.container(border=True):
        st.markdown("### Período de Referência")
        c_p1, c_p2, _ = st.columns(3)
        with c_p1: st.selectbox("Mês", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"], key="sel_mes")
        with c_p2: st.selectbox("Ano", [2024, 2025, 2026, 2027], index=2, key="sel_ano")
    
    with st.container(border=True):
        st.markdown("### Atendimentos Ambulatoriais")
        c1, c2, c3 = st.columns(3)
        with c1: st.number_input("Clínica Médica", key="in_h_cli_med", step=1)
        with c2: st.number_input("Ortopedista", key="in_h_orto", step=1)
        with c3: st.number_input("Cardiologia", key="in_h_card", step=1)
        c4, c5, c6 = st.columns(3)
        with c4: st.number_input("Neurologista", key="in_h_neuro", step=1)
        with c5: st.number_input("Pediatria", key="in_h_ped", step=1)
        with c6: st.number_input("Ginecologista", key="in_h_gineco", step=1)
        c7, c8, c9 = st.columns(3)
        with c7: st.number_input("Psiquiatra", key="in_h_psiq", step=1)
        with c8: st.number_input("Gastroenterologista", key="in_h_gastr", step=1)
        with c9: st.number_input("Cirurgião Geral", key="in_h_cir_gr", step=1)
    
    with st.container(border=True):
        st.markdown("### Consultas Não Médicas")
        c10, c11, c12 = st.columns(3)
        with c10: st.number_input("Psicólogo", key="in_h_psico", step=1)
        with c11: st.number_input("Psicopedagogo", key="in_h_psic_ped", step=1)
        with c12: st.number_input("Fonoaudiólogo", key="in_h_fono", step=1)
        with st.columns(3)[0]: st.number_input("Terapeuta Ocupacional", key="in_h_terap", step=1)

    with st.container(border=True):
        st.markdown("### Cirurgias e Internações")
        c15, c16, c17 = st.columns(3)
        with c15: st.number_input("Total Proc. Cirúrgicos", key="in_h_t_cirurgia", step=1)
        with c16: st.number_input("Total Cirurgia Geral", key="in_h_t_cir_gr", step=1)
        with c17: st.number_input("Total Cirurgia Gineco/Obst", key="in_h_t_cir_gin", step=1)
        c18, c19, c20 = st.columns(3)
        with c18: st.number_input("Total Pacientes Internados", key="in_h_t_pac_int", step=1)
        with c19: st.number_input("Saída por Alta", key="in_h_s_alta", step=1)
       

    with st.container(border=True):
        st.markdown("### Óbitos e Permanência")
        c21, c22, _ = st.columns(3)
        with c21: st.number_input("Saída Óbito > 24H", key="in_h_ob_maior", step=1)
        with c22: st.number_input("Saída Óbito < 24H", key="in_h_ob_menor", step=1)
        c1, c2, _= st.columns(3)
        with c1: st.number_input("Permanência < 24H", key="in_h_temp_perm_menor", step=1)
        with c2: st.number_input("Permanência > 24H", key="in_h_temp_perm_maior", step=1)

    with st.container(border=True):
        st.markdown("### Saídas por Clínica")
        c24, c25, _ = st.columns(3)
        with c24: st.number_input("Saída Clínica Médica", key="in_h_s_climed", step=1)
        with c25: st.number_input("Saída Clínica Cirúrgica", key="in_h_s_clicir", step=1)
        c1, c2, _ = st.columns(3)
        with c1: st.number_input("Saída Clínica Obstétrica", key="in_h_s_cliobs", step=1)
        with c2: st.number_input("Saída Clínica Pediátrica", key="in_h_s_cliped", step=1)

    with st.container(border=True):
        st.markdown("### Emergência")
        with st.columns(3)[0]: st.number_input("Total Pacientes Emergência", key="in_total_paci_emerg", step=1)

# --- ABA CAPS ---
with t_manual_caps:
    with st.container(border=True):
        st.markdown("### Indicadores CAPS")
        c1, c2 = st.columns(2)
        with c1: st.number_input("Total Atendimentos CAPS", key="in_caps_t_atend", step=1)
        with c2: st.number_input("Atendimento Individual", key="in_caps_atend_ind", step=1)
        c1, c2 = st.columns(2)
        with c1: st.number_input("Atendimento de Grupo", key="in_caps_atend_grp", step=1)
        with c2: st.number_input("Quantidade de Grupos CAPS", key="in_caps_t_grupos", step=1)

# --- ABA ATENÇÃO BÁSICA ---
with t_manual_ab:
    with st.container(border=True):
        st.markdown("### Indicadores AB")
        ab1, ab2 = st.columns(2)
        with ab1: st.number_input("Consultas Médicas AB", key="in_ab_cons_med", step=1)
        with ab2: st.number_input("Consultas Enfermagem AB", key="in_ab_cons_enf", step=1)
        ab1, ab2 = st.columns(2)
        with ab1: st.number_input("Atendimento Odonto AB", key="in_ab_atend_odont", step=1)
        with ab2: st.number_input("Visita Domiciliar AB", key="in_ab_vist_domi", step=1)

# --- ABA ARQUIVOS (EVIDÊNCIAS) ---
with t_evidencia:
    labels = {
        "H_GRAFICO_ATEND_AMB": "Gráfico de Atendimento Hospitalar",
        "H_GRAFICO_ASSIST_HOSP": "Gráfico de Assistência Hospitalar",
        "H_TAB_TRANS_HOSP": "Tabela de Transferência Hospitalar",
        "H_GRAFICO_TRANS_HOSP": "Gráfico de Transferência Hospitalar",
        "H_GRAFICO_SAIDA_HOSP": "Gráfico de Saída Hospitalar",
        "H_GRAFICO_ATEND_EMERG": "Gráfico de Atendimento de Emergência",
        "ATA_REUNIAO_OBITO": "Ata Revisão de Óbito",
        "ATA_REUNIAO_PRONTUARIO": "Ata Revisão de Prontuário",
        "ATA_REUNIAO_INFEC": "Ata Controle de Infecção",
        "CAPS_GRAFICO_ATEND": "Gráfico de Atendimento CAPS",
        "CAPS_REGISTRO_FOTOGRAFICO": "Registro Fotográfico CAPS",
        "AB_GRAFICO_ATEND": "Gráfico de Atendimento Antenção Básica",
        "AB_METAQUANTI_HOSP": "Tabela Quanti",
        "AB_METAQUALI_HOSP": "Tabela Quali"
        }
    grupos_evidencias = [
        {"nome": "Ambulatorial / Hospitalar", "marcadores": ["H_GRAFICO_ATEND_AMB", "H_GRAFICO_ASSIST_HOSP", "H_TAB_TRANS_HOSP", "H_GRAFICO_TRANS_HOSP", "H_GRAFICO_SAIDA_HOSP"]},
        {"nome": "Emergência e Atas", "marcadores": ["H_GRAFICO_ATEND_EMERG", "ATA_REUNIAO_OBITO", "ATA_REUNIAO_PRONTUARIO", "ATA_REUNIAO_INFEC"]},
        {"nome": "CAPS", "marcadores": ["CAPS_GRAFICO_ATEND", "CAPS_REGISTRO_FOTOGRAFICO"]},
        {"nome": "Atenção Básica", "marcadores": ["AB_GRAFICO_ATEND", "AB_METAQUANTI_HOSP", "AB_METAQUALI_HOSP"]}
    ]

    for grupo in grupos_evidencias:
        with st.container(border=True):
            st.subheader(grupo["nome"])
            ce1, ce2 = st.columns(2)
            for idx, m in enumerate(grupo["marcadores"]):
                target = ce1 if idx % 2 == 0 else ce2
                with target:
                    st.markdown(f"<span class='upload-label'>{labels.get(m, m)}</span>", unsafe_allow_html=True)
                    f_up = st.file_uploader("Upload", type=['png', 'jpg', 'pdf'], key=f"f_{m}", label_visibility="collapsed")
                    if f_up:
                        if f_up.name not in [x['name'] for x in st.session_state.dados_sessao.get(m, [])]:
                            st.session_state.dados_sessao[m].append({"name": f_up.name, "content": f_up, "type": "f"})
                            st.rerun()

                    kp = f"p_{m}_{len(st.session_state.dados_sessao.get(m, []))}"
                    pasted = paste_image_button(label="📸 Colar Print", key=kp)
                    if pasted is not None and pasted.image_data is not None:
                        st.session_state.dados_sessao[m].append({"name": f"Captura_{m}.png", "content": pasted.image_data, "type": "p"})
                        st.toast(f"Anexado: {labels.get(m)}")
                        time.sleep(0.4); st.rerun()
                                        
                    if st.session_state.dados_sessao.get(m):
                        for i_idx, item in enumerate(st.session_state.dados_sessao[m]):
                            with st.expander(f"📄 {item['name']}", expanded=False):
                                is_img = item['type'] == "p" or item['name'].lower().endswith(('.png', '.jpg', '.jpeg'))
                                if is_img: st.image(item['content'], width='stretch')
                                else: st.info(f"PDF pronto.")
                                if st.button("Remover", key=f"del_{m}_{i_idx}"):
                                    st.session_state.dados_sessao[m].pop(i_idx)
                                    st.rerun()

# --- GERAÇÃO FINAL ---
if st.button("FINALIZAR E GERAR RELATÓRIO", type="primary", width='stretch'):
    try:
        with st.spinner("Gerando documentos..."):
            with tempfile.TemporaryDirectory() as tmp:
                docx_p = os.path.join(tmp, "relatorio.docx")
                doc = DocxTemplate("template-madalena.docx")
                
                h_atend_esp_med = sum([int(st.session_state.get(k, 0) or 0) for k in ["in_h_cli_med", "in_h_orto", "in_h_card", "in_h_neuro", "in_h_ped", "in_h_gineco", "in_h_psiq", "in_h_gastr", "in_h_cir_gr"]])
                h_consulta_nao_med = sum([int(st.session_state.get(k, 0) or 0) for k in ["in_h_psico", "in_h_psic_ped", "in_h_fono", "in_h_terap"]])
                total_amb = h_atend_esp_med + h_consulta_nao_med
                
                total_saidas = sum([int(st.session_state.get(k, 0) or 0) for k in ["in_h_s_climed", "in_h_s_clicir", "in_h_s_cliobs", "in_h_s_cliped"]])
                total_obitos = int(st.session_state.get("in_h_ob_maior", 0) or 0) + int(st.session_state.get("in_h_ob_menor", 0) or 0)
                total_ab = sum([int(st.session_state.get(k, 0) or 0) for k in ["in_ab_cons_med", "in_ab_cons_enf", "in_ab_atend_odont", "in_ab_vist_domi"]])
                h_s_trans = sum([int(st.session_state.get(k, 0) or 0) for k in ["in_h_temp_perm_menor", "in_h_temp_perm_maior"]])

                dados_finais = {
                    "SISTEMA_MES_REFERENCIA": f"{st.session_state.get('sel_mes', 'Janeiro')}/{st.session_state.get('sel_ano', 2026)}",
                    "H_TOTAL_ATEND_AMB": total_amb,
                    "H_ATEND_ESP_MED": h_atend_esp_med,
                    "H_CONSULTA_NAO_MED": h_consulta_nao_med,
                    "H_CLI_MED": st.session_state.get("in_h_cli_med", 0),
                    "H_ORTO": st.session_state.get("in_h_orto", 0),
                    "H_CARD": st.session_state.get("in_h_card", 0),
                    "H_NEURO": st.session_state.get("in_h_neuro", 0),
                    "H_PED": st.session_state.get("in_h_ped", 0),
                    "H_GINECO": st.session_state.get("in_h_gineco", 0),
                    "H_PSIQ": st.session_state.get("in_h_psiq", 0),
                    "H_GASTR": st.session_state.get("in_h_gastr", 0),
                    "H_CIR_GR": st.session_state.get("in_h_cir_gr", 0),
                    "H_PSICO": st.session_state.get("in_h_psico", 0),
                    "H_PSIC_PED": st.session_state.get("in_h_psic_ped", 0),
                    "H_FONO": st.session_state.get("in_h_fono", 0),
                    "H_TERAP": st.session_state.get("in_h_terap", 0),
                    "H_T_CIRURGIA": st.session_state.get("in_h_t_cirurgia", 0),
                    "H_T_CIR_GR": st.session_state.get("in_h_t_cir_gr", 0),
                    "H_T_CIR_GIN": st.session_state.get("in_h_t_cir_gin", 0),
                    "H_T_PAC_INT": st.session_state.get("in_h_t_pac_int", 0),
                    "H_S_ALTA": st.session_state.get("in_h_s_alta", 0),
                    "H_S_TRANS": h_s_trans,
                    "H_OB_MAIOR": st.session_state.get("in_h_ob_maior", 0),
                    "H_OB_MENOR": st.session_state.get("in_h_ob_menor", 0),
                    "H_TEMP_PERM_MENOR": st.session_state.get("in_h_temp_perm_menor", 0),
                    "H_TEMP_PERM_MAIOR": st.session_state.get("in_h_temp_perm_maior", 0),
                    "H_S_CLIMED": st.session_state.get("in_h_s_climed", 0),
                    "H_S_CLICIR": st.session_state.get("in_h_s_clicir", 0),
                    "H_S_CLIOBS": st.session_state.get("in_h_s_cliobs", 0),
                    "H_S_CLIPED": st.session_state.get("in_h_s_cliped", 0),
                    "TOTAL_PACI_EMERG": st.session_state.get("in_total_paci_emerg", 0),
                    "H_T_SAIDA": total_saidas,
                    "H_TOTAL_OBITO": total_obitos,
                    "CAPS_T_ATEND": st.session_state.get("in_caps_t_atend", 0),
                    "CAPS_ATEND_IND": st.session_state.get("in_caps_atend_ind", 0),
                    "CAPS_ATEND_GRP": st.session_state.get("in_caps_atend_grp", 0),
                    "CAPS_T_GRUPOS": st.session_state.get("in_caps_t_grupos", 0),
                    "AB_CONS_MED": st.session_state.get("in_ab_cons_med", 0),
                    "AB_CONS_ENF": st.session_state.get("in_ab_cons_enf", 0),
                    "AB_ATEND_ODONT": st.session_state.get("in_ab_atend_odont", 0),
                    "AB_VIST_DOMI": st.session_state.get("in_ab_vist_domi", 0),
                    "AB_T_ATEND": total_ab
                }

                for marcador in DIMENSOES_CAMPOS.keys():
                    imgs_word = []
                    for item in st.session_state.dados_sessao.get(marcador, []):
                        res = processar_item_lista(doc, item['content'], marcador)
                        if res: imgs_word.extend(res)
                    dados_finais[marcador] = imgs_word
                
                try:
                    doc.render(dados_finais)
                except Exception as jinja_err:
                    st.error(f"Erro de Sintaxe no Word: {jinja_err}")
                    st.stop()
                
                doc.save(docx_p)
                st.success("✅ Relatório gerado com sucesso!")
                cd1, cd2 = st.columns(2)
                with cd1:
                    with open(docx_p, "rb") as f_w:
                        st.download_button("WORD (.docx)", f_w.read(), f"RELATÓRIO ASSISTENCIAL MENSAL - SANTA MARIA MADALENA {st.session_state.get('sel_mes', 'MAD')}.docx", width='stretch')
                with cd2:
                    try:
                        converter_para_pdf(docx_p, tmp)
                        pdf_p = os.path.join(tmp, "relatorio.pdf")
                        if os.path.exists(pdf_p):
                            with open(pdf_p, "rb") as f_p:
                                st.download_button("PDF", f_p.read(), f"RELATÓRIO ASSISTENCIAL MENSAL - SANTA MARIA MADALENA {st.session_state.get('sel_mes', 'MAD')}.pdf", width='stretch')
                    except: st.warning("PDF falhou.")
    except Exception as e: st.error(f"Erro Crítico: {e}")

st.caption("Desenvolvido por Leonardo Barcelos Martins")

