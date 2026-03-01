# -*- coding: utf-8 -*-
import os
import re
import io
import json
import time
import zipfile
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
from streamlit_autorefresh import st_autorefresh
import smtplib
from email.message import EmailMessage

st.set_page_config(page_title="Aplicativo de Questões - TOMOs", layout="wide")

# Dependências opcionais
try:
    from docx import Document
    DOCX_DISPONIVEL = True
    DOCX_IMPORT_ERROR = None
except Exception as e:
    DOCX_DISPONIVEL = False
    DOCX_IMPORT_ERROR = str(e)

try:
    from PIL import Image
    PIL_DISPONIVEL = True
except ImportError:
    PIL_DISPONIVEL = False


# ============================
# CONFIG
# ============================
ARQUIVO_PROGRESSO = "progresso_questoes.json"
DIRETORIO_RELATORIOS = "relatorios"
TEMPO_PADRAO_MIN = 60
ARQUIVO_PLANILHA_PADRAO = "ADS_POO.xlsx"


# ============================
# ESTADO INICIAL
# ============================
def init_session():
    defaults = {
        "df": None,
        "questoes_agrupadas": [],
        "questao_atual": 0,
        "alternativa_selecionada": None,
        "resposta_verificada": False,

        # estatísticas
        "total_acertos": 0,
        "total_erros": 0,
        "questoes_respondidas": set(),
        "respostas_por_questao": {},  # idx -> letra (congelada no verificar)

        # usuário
        "nome_usuario": "",
        "registro_academico": "",
        "turma": "",

        # modo
        "modo_atual": "estudo",
        "tempo_prova": 60 * 60,
        "tempo_restante": None,
        "timer_ativo": False,
        "prova_inicio_epoch": None,
        "prova_fim_epoch": None,
        "questoes_iniciadas": False,

        # diretórios
        "diretorio_teorias": "teorias",
        "diretorio_imagens": "imagens",

        # flags
        "cadastro_confirmado": False,
        "planilha_carregada": False,
        "mostrar_justificativas": False,
        "status_msg": "",

        # usado para recriar widgets radio quando precisar (ex.: recarregar planilha)
        "radio_reset_version": 0,

        # controle de abertura da teoria
        "abrir_teoria_para_idx": None,
    }

    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ============================
# UTILITÁRIOS
# ============================
def to_int_safe(v, default=0):
    try:
        return int(float(v))
    except Exception:
        return default


def eh_algarismo_romano(texto):
    texto = str(texto).strip().upper()
    padrao_romano = r'^[IVXLCDM]+$'
    return bool(re.match(padrao_romano, texto)) and len(texto) <= 5


def salvar_json(caminho, dados):
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)
def enviar_relatorio_por_email(txt_content: str, json_content: str, nome_txt: str, nome_json: str):
    """
    Envia relatório por e-mail usando SMTP e credenciais em st.secrets.
    Espera existir st.secrets["email"] com as chaves configuradas.
    """
    try:
        email_cfg = st.secrets["email"]

        smtp_server = email_cfg["smtp_server"]
        smtp_port = int(email_cfg["smtp_port"])
        username = email_cfg["username"]
        password = email_cfg["password"]
        from_email = email_cfg.get("from_email", username)
        to_email = email_cfg.get("to_email", "pedro.euphrasio@docente.unip.br")
        use_tls = bool(email_cfg.get("use_tls", True))

        # Monta mensagem
        msg = EmailMessage()
        msg["Subject"] = (
            f"Relatório de Avaliação - {st.session_state.nome_usuario} "
            f"(RA {st.session_state.registro_academico}) - "
            f"{datetime.now().strftime('%d/%m/%Y %H:%M')}"
        )
        msg["From"] = from_email
        msg["To"] = to_email

        corpo = f"""
Olá,

Segue em anexo o relatório de avaliação gerado pelo aplicativo.

Aluno: {st.session_state.nome_usuario}
RA: {st.session_state.registro_academico}
Turma: {st.session_state.turma}
Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}

Atenciosamente,
Aplicativo de Questões (Streamlit)
""".strip()

        msg.set_content(corpo)

        # Anexos
        msg.add_attachment(
            txt_content.encode("utf-8"),
            maintype="text",
            subtype="plain",
            filename=nome_txt
        )

        msg.add_attachment(
            json_content.encode("utf-8"),
            maintype="application",
            subtype="json",
            filename=nome_json
        )

        # Envio SMTP
        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
            if use_tls:
                server.starttls()
            server.login(username, password)
            server.send_message(msg)

        return True, f"✅ Relatório enviado por e-mail para {to_email}"

    except KeyError as e:
        return False, f"❌ Configuração ausente em st.secrets[email]: {e}"
    except Exception as e:
        return False, f"❌ Erro ao enviar e-mail: {e}"

def carregar_json(caminho):
    if not os.path.exists(caminho):
        return None
    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)
def formatar_tempo(segundos):
    try:
        segundos = 0 if segundos is None else int(float(segundos))
        segundos = max(0, segundos)
        return f"{segundos // 60:02d}:{segundos % 60:02d}"
    except Exception:
        return "00:00"
def atualizar_tempo_prova():
    if st.session_state.modo_atual != "prova":
        return

    if not st.session_state.timer_ativo:
        return

    fim = st.session_state.get("prova_fim_epoch", None)
    if fim is None:
        return

    restante = int(fim - time.time())

    if restante <= 0:
        st.session_state.tempo_restante = 0
        st.session_state.timer_ativo = False
    else:
        st.session_state.tempo_restante = restante


def calcular_tempo_utilizado():
    if st.session_state.modo_atual == "prova":
        if st.session_state.tempo_restante is None:
            return formatar_tempo(st.session_state.tempo_prova)
        tempo_usado = st.session_state.tempo_prova - st.session_state.tempo_restante
        return formatar_tempo(tempo_usado)
    return "00:00"


# ============================
# PROGRESSO
# ============================
def salvar_progresso():
    progresso = {
        "data_salvamento": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_acertos": st.session_state.total_acertos,
        "total_erros": st.session_state.total_erros,
        "questoes_respondidas": list(st.session_state.questoes_respondidas),
        "questao_atual": st.session_state.questao_atual,
        "modo_atual": st.session_state.modo_atual,
        "tempo_prova": st.session_state.tempo_prova,
        "tempo_restante": st.session_state.tempo_restante,
        "diretorio_teorias": st.session_state.diretorio_teorias,
        "diretorio_imagens": st.session_state.diretorio_imagens,
        "dados_usuario": {
            "nome": st.session_state.nome_usuario,
            "registro_academico": st.session_state.registro_academico,
            "turma": st.session_state.turma
        },
        "respostas_por_questao": st.session_state.respostas_por_questao,
        "cadastro_confirmado": st.session_state.cadastro_confirmado,
        "radio_reset_version": st.session_state.radio_reset_version,
    }
    try:
        salvar_json(ARQUIVO_PROGRESSO, progresso)
        st.session_state.status_msg = f"💾 Progresso salvo ({datetime.now().strftime('%H:%M:%S')})"
    except Exception as e:
        st.session_state.status_msg = f"❌ Erro ao salvar progresso: {e}"


def carregar_progresso():
    progresso = carregar_json(ARQUIVO_PROGRESSO)
    if not progresso:
        return

    try:
        st.session_state.total_acertos = progresso.get("total_acertos", 0)
        st.session_state.total_erros = progresso.get("total_erros", 0)
        st.session_state.questoes_respondidas = set(progresso.get("questoes_respondidas", []))
        st.session_state.questao_atual = progresso.get("questao_atual", 0)
        st.session_state.modo_atual = progresso.get("modo_atual", "estudo")
        st.session_state.tempo_prova = progresso.get("tempo_prova", TEMPO_PADRAO_MIN * 60)
        st.session_state.tempo_restante = progresso.get("tempo_restante", None)
        st.session_state.diretorio_teorias = progresso.get("diretorio_teorias", "teorias")
        st.session_state.diretorio_imagens = progresso.get("diretorio_imagens", "imagens")
        st.session_state.respostas_por_questao = progresso.get("respostas_por_questao", {})
        st.session_state.cadastro_confirmado = progresso.get("cadastro_confirmado", False)
        st.session_state.radio_reset_version = progresso.get("radio_reset_version", 0)

        dados_usuario = progresso.get("dados_usuario", {})
        st.session_state.nome_usuario = dados_usuario.get("nome", "")
        st.session_state.registro_academico = dados_usuario.get("registro_academico", "")
        st.session_state.turma = dados_usuario.get("turma", "")

        st.session_state.status_msg = f"✅ Progresso carregado ({progresso.get('data_salvamento', 'data desconhecida')})"
    except Exception as e:
        st.session_state.status_msg = f"⚠️ Erro ao carregar progresso: {e}"


# ============================
# PLANILHA / QUESTÕES
# ============================
def agrupar_questoes_corretamente(df):
    questoes = []
    questao_atual = None

    for _, row in df.iterrows():
        q_col = "Questão "
        if q_col not in df.columns:
            raise KeyError("Coluna 'Questão ' não encontrada na planilha.")

        if pd.notna(row.get("Questão ")):
            if questao_atual is not None:
                questoes.append(questao_atual)

            questao_atual = {
                "tomo": row.get("TOMO", "N/A") if pd.notna(row.get("TOMO", "N/A")) else "N/A",
                "numero": row.get("Questão "),
                "enunciado": row.get("Enunciado", "") if pd.notna(row.get("Enunciado", "")) else "",
                "imagem": row.get("Imagem", "") if pd.notna(row.get("Imagem", "")) else "",
                "alternativas": [],
                "usa_algarismos_romanos": False,
            }

        if pd.notna(row.get("Alternativas")) and questao_atual is not None:
            alt_letra = str(row.get("Alternativas")).strip()
            alternativa = {
                "letra": alt_letra,
                "texto": str(row.get("Textos das Alternativas", "")).strip() if pd.notna(row.get("Textos das Alternativas")) else "",
                "analise": str(row.get("Análise das alternativas ", "")).strip() if pd.notna(row.get("Análise das alternativas ")) else "",
                "correta": str(row.get("Alternativa correta", "")).strip() == "X" if pd.notna(row.get("Alternativa correta")) else False
            }
            questao_atual["alternativas"].append(alternativa)

            if eh_algarismo_romano(alt_letra):
                questao_atual["usa_algarismos_romanos"] = True

    if questao_atual is not None:
        questoes.append(questao_atual)

    return questoes


def carregar_planilha_arquivo(uploaded_file=None, caminho_local=None):
    if uploaded_file is None and not caminho_local:
        st.error("Informe um arquivo (upload) ou caminho local.")
        return False

    try:
        if uploaded_file is not None:
            nome = uploaded_file.name.lower()
            bytes_data = uploaded_file.read()
            bio = io.BytesIO(bytes_data)
            if nome.endswith(".xlsx") or nome.endswith(".xlsm"):
                df = pd.read_excel(bio, engine="openpyxl")
            elif nome.endswith(".xls"):
                df = pd.read_excel(bio, engine="xlrd")
            else:
                df = pd.read_excel(bio)
        else:
            if not os.path.exists(caminho_local):
                st.error(f"Arquivo não encontrado: {caminho_local}")
                return False

            ext = os.path.splitext(caminho_local)[1].lower()
            if ext in [".xlsx", ".xlsm"]:
                df = pd.read_excel(caminho_local, engine="openpyxl")
            elif ext == ".xls":
                df = pd.read_excel(caminho_local, engine="xlrd")
            else:
                df = pd.read_excel(caminho_local)

        questoes = agrupar_questoes_corretamente(df)
        if not questoes:
            st.warning("Nenhuma questão encontrada na planilha.")
            return False

        st.session_state.df = df
        st.session_state.questoes_agrupadas = questoes
        st.session_state.planilha_carregada = True

        # reset estatísticas ao carregar nova planilha
        st.session_state.total_acertos = 0
        st.session_state.total_erros = 0
        st.session_state.questoes_respondidas = set()
        st.session_state.respostas_por_questao = {}
        st.session_state.questao_atual = 0
        st.session_state.resposta_verificada = False
        st.session_state.alternativa_selecionada = None
        st.session_state.mostrar_justificativas = False

        # força recriação dos radios
        st.session_state.radio_reset_version += 1

        st.success(f"✅ {len(questoes)} questões carregadas com sucesso!")
        salvar_progresso()
        return True

    except ImportError as e:
        if "xlrd" in str(e):
            st.error("Biblioteca 'xlrd' não encontrada. Instale com: pip install xlrd")
        elif "openpyxl" in str(e):
            st.error("Biblioteca 'openpyxl' não encontrada. Instale com: pip install openpyxl")
        else:
            st.error(f"Erro de importação: {e}")
        return False
    except Exception as e:
        st.error(f"Erro ao carregar planilha: {e}")
        return False


# ============================
# IMAGENS DAS QUESTÕES
# ============================
def encontrar_imagem_questao(nome_imagem, diretorio_imagens):
    if not nome_imagem:
        return None

    if not os.path.exists(diretorio_imagens):
        return None

    extensoes = [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"]
    for ext in extensoes:
        p = os.path.join(diretorio_imagens, str(nome_imagem) + ext)
        if os.path.exists(p):
            return p

    p_exato = os.path.join(diretorio_imagens, str(nome_imagem))
    if os.path.exists(p_exato):
        return p_exato

    return None


def mostrar_imagem_questao_streamlit(nome_imagem):
    caminho = encontrar_imagem_questao(nome_imagem, st.session_state.diretorio_imagens)
    if not caminho:
        st.info(f"🖼️ Imagem não encontrada para: {nome_imagem}")
        return

    try:
        if PIL_DISPONIVEL:
            img = Image.open(caminho)
            st.image(img, caption=f"Figura: {nome_imagem}", width="stretch")
        else:
            st.image(caminho, caption=f"Figura: {nome_imagem}", width="stretch")
    except Exception as e:
        st.warning(f"Erro ao exibir imagem {nome_imagem}: {e}")


# ============================
# TEORIAS (.docx)
# ============================
def localizar_arquivo_teoria(questao):
    if not os.path.exists(st.session_state.diretorio_teorias):
        return None, []

    tomo = to_int_safe(questao["tomo"], 0)
    numero = to_int_safe(questao["numero"], 0)

    possiveis_nomes = [
        f"T{tomo}Q{numero}.docx",
        f"T{tomo}Q{numero:02d}.docx",
        f"Tomo{tomo}Questao{numero}.docx",
        f"Tomo {tomo} Questão {numero}.docx",
        f"Tomo{tomo}Q{numero}.docx",
    ]

    for nome in possiveis_nomes:
        caminho = os.path.join(st.session_state.diretorio_teorias, nome)
        if os.path.exists(caminho):
            return caminho, possiveis_nomes

    return None, possiveis_nomes


def extrair_imagens_docx(caminho_docx):
    imagens = []
    try:
        with zipfile.ZipFile(caminho_docx) as z:
            media_files = [f for f in z.namelist() if f.startswith("word/media/")]
            for mf in media_files:
                try:
                    data = z.read(mf)
                    imagens.append((os.path.basename(mf), data))
                except Exception:
                    continue
    except Exception:
        pass
    return imagens


def renderizar_teoria(questao):
    st.subheader("📖 Teoria da questão")

    if not DOCX_DISPONIVEL:
        st.error("Biblioteca python-docx não encontrada. Instale com: pip install python-docx")
        return

    caminho, possiveis = localizar_arquivo_teoria(questao)
    if not caminho:
        st.warning("Arquivo de teoria não encontrado.")
        with st.expander("Ver nomes esperados"):
            st.write("Formatos procurados:")
            for n in possiveis:
                st.write(f"- {n}")
            st.write(f"Diretório atual: `{st.session_state.diretorio_teorias}`")
        return

    st.caption(f"Arquivo: {os.path.basename(caminho)}")

    try:
        doc = Document(caminho)

        for p in doc.paragraphs:
            txt = p.text.strip()
            if not txt:
                continue

            style_name = getattr(p.style, "name", "") if p.style else ""
            if str(style_name).startswith("Heading"):
                st.markdown(f"### {txt}")
            else:
                st.write(txt)

        imagens = extrair_imagens_docx(caminho)
        if imagens:
            st.markdown("#### 🖼 Imagens da teoria")
            for nome_img, data in imagens:
                try:
                    if PIL_DISPONIVEL:
                        img = Image.open(BytesIO(data))
                        st.image(img, caption=nome_img, width="stretch")
                    else:
                        st.image(data, caption=nome_img, width="stretch")
                except Exception:
                    continue

    except Exception as e:
        st.error(f"Erro ao ler teoria: {e}")


# ============================
# RESPOSTAS / JUSTIFICATIVAS
# ============================
def get_alt_correta(questao):
    for alt in questao["alternativas"]:
        if alt["correta"]:
            return alt["letra"]
    return None


def limpar_analise(analise):
    if not analise:
        return "Justificativa não disponível."
    a = analise.strip()
    if a.startswith("Alternativa correta."):
        a = a.replace("Alternativa correta.", "", 1).strip()
    elif a.startswith("Alternativa incorreta."):
        a = a.replace("Alternativa incorreta.", "", 1).strip()
    return a or "Justificativa não disponível."


def verificar_resposta_streamlit():
    idx = st.session_state.questao_atual

    if idx in st.session_state.questoes_respondidas:
        st.warning("🔒 Esta questão já foi verificada. A resposta está bloqueada.")
        return

    questao = st.session_state.questoes_agrupadas[idx]
    key_radio = f"resp_q_{idx}_v{st.session_state.radio_reset_version}"
    alt_sel = st.session_state.get(key_radio, "")

    if not alt_sel:
        st.warning("Por favor, selecione uma alternativa!")
        return

    alt_correta = get_alt_correta(questao)
    acertou = (alt_sel == alt_correta)

    if acertou:
        st.session_state.total_acertos += 1
    else:
        st.session_state.total_erros += 1

    st.session_state.questoes_respondidas.add(idx)
    st.session_state.respostas_por_questao[str(idx)] = alt_sel  # congela resposta

    st.session_state.alternativa_selecionada = alt_sel
    st.session_state.resposta_verificada = True
    st.session_state.mostrar_justificativas = False

    salvar_progresso()


def renderizar_feedback(questao, alt_sel, idx=None):
    """
    Feedback estilo prova:
    - Após verificar, mostra que foi registrada e bloqueada
    - Não revela a alternativa correta aqui
    """
    if idx is None:
        idx = st.session_state.questao_atual

    alt_correta = get_alt_correta(questao)

    # se já verificada, usa congelada
    if idx in st.session_state.questoes_respondidas:
        resposta_congelada = st.session_state.respostas_por_questao.get(str(idx), alt_sel)
        if not resposta_congelada:
            st.info("🔒 Questão já verificada. A resposta está bloqueada.")
            return

        if resposta_congelada == alt_correta:
            st.success(f"✅ Questão verificada. Resposta registrada: {resposta_congelada}.")
        else:
            st.warning(
                f"🔒 Questão verificada. Resposta registrada: {resposta_congelada}. "
                f"Não é possível alterar após clicar em 'Verificar'."
            )
        return

    # ainda não verificada
    acertou = (alt_sel == alt_correta)
    if acertou:
        st.success(f"🎉 Resposta correta! Você marcou **{alt_sel}**.")
    else:
        st.error(f"❌ Resposta incorreta. Você marcou **{alt_sel}**.")
        st.info("Clique em **Mostrar Resposta / Justificativas** para ver a alternativa correta e as explicações.")


def renderizar_justificativas(questao):
    alt_correta = get_alt_correta(questao)
    idx = st.session_state.questao_atual

    # se já verificada, a "sua resposta" deve ser a congelada
    alt_sel = st.session_state.respostas_por_questao.get(str(idx), st.session_state.alternativa_selecionada)

    st.markdown("## 📝 Resposta e Justificativas")
    if alt_correta:
        st.success(f"✅ Alternativa correta: **{alt_correta}**")
    if alt_sel:
        if alt_sel == alt_correta:
            st.info(f"✅ Sua resposta: **{alt_sel}**")
        else:
            st.info(f"❌ Sua resposta: **{alt_sel}**")

    if questao["usa_algarismos_romanos"]:
        st.markdown("### 📚 Justificativas de todas as alternativas (algarismos romanos)")
        for alt in questao["alternativas"]:
            cor = "🟢" if alt["correta"] else "🔴"
            with st.expander(f"{cor} Alternativa {alt['letra']}"):
                if alt["texto"]:
                    st.caption(f"Texto: {alt['texto']}")
                st.write(limpar_analise(alt["analise"]))
    else:
        st.markdown("### 📚 Clique na alternativa para ver a justificativa")
        tabs = st.tabs([f"Alt {a['letra']}" for a in questao["alternativas"]])
        for tab, alt in zip(tabs, questao["alternativas"]):
            with tab:
                if alt["correta"]:
                    st.success(f"✅ Alternativa {alt['letra']}")
                else:
                    st.error(f"❌ Alternativa {alt['letra']}")
                if alt["texto"]:
                    st.caption(f"Texto: {alt['texto']}")
                st.write(limpar_analise(alt["analise"]))


def gerar_relatorio():
    if not st.session_state.questoes_agrupadas:
        st.warning("Nenhuma questão foi carregada ainda!")
        return False, None

    total_questoes = len(st.session_state.questoes_agrupadas)
    total_respondidas = len(st.session_state.questoes_respondidas)
    aproveitamento = (st.session_state.total_acertos / total_respondidas * 100) if total_respondidas > 0 else 0

    relatorio = {
        "dados_usuario": {
            "nome": st.session_state.nome_usuario,
            "registro_academico": st.session_state.registro_academico,
            "turma": st.session_state.turma,
            "data_cadastro": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        },
        "estatisticas": {
            "total_questoes": total_questoes,
            "questoes_respondidas": total_respondidas,
            "acertos": st.session_state.total_acertos,
            "erros": st.session_state.total_erros,
            "aproveitamento": f"{aproveitamento:.1f}%",
            "data_geracao": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        },
        "detalhes_questoes": []
    }

    for idx in sorted(list(st.session_state.questoes_respondidas)):
        if idx < len(st.session_state.questoes_agrupadas):
            q = st.session_state.questoes_agrupadas[idx]
            relatorio["detalhes_questoes"].append({
                "tomo": q["tomo"],
                "numero": to_int_safe(q["numero"], 0),
                "alternativa_correta": get_alt_correta(q),
                "alternativa_marcada": st.session_state.respostas_por_questao.get(str(idx), None)
            })

    try:
        # -----------------------------
        # Conteúdo em memória (JSON/TXT)
        # -----------------------------
        json_content = json.dumps(relatorio, indent=4, ensure_ascii=False)

        linhas_txt = []
        linhas_txt.append("=" * 60)
        linhas_txt.append("RELATÓRIO DE DESEMPENHO - APLICATIVO DE QUESTÕES")
        linhas_txt.append("=" * 60)
        linhas_txt.append("")
        linhas_txt.append("DADOS DO ESTUDANTE:")
        linhas_txt.append("-" * 60)
        linhas_txt.append(f"Nome: {st.session_state.nome_usuario}")
        linhas_txt.append(f"Registro Acadêmico: {st.session_state.registro_academico}")
        linhas_txt.append(f"Turma: {st.session_state.turma}")
        linhas_txt.append(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        linhas_txt.append("")
        linhas_txt.append("ESTATÍSTICAS:")
        linhas_txt.append("-" * 60)
        linhas_txt.append(f"Total de Questões: {total_questoes}")
        linhas_txt.append(f"Questões Respondidas: {total_respondidas}")
        linhas_txt.append(f"Acertos: {st.session_state.total_acertos}")
        linhas_txt.append(f"Erros: {st.session_state.total_erros}")
        linhas_txt.append(f"Aproveitamento: {aproveitamento:.1f}%")
        linhas_txt.append(f"Tempo utilizado: {calcular_tempo_utilizado()}")
        linhas_txt.append("")
        linhas_txt.append("DETALHES DAS QUESTÕES:")
        linhas_txt.append("-" * 60)

        for d in relatorio["detalhes_questoes"]:
            linhas_txt.append(
                f"TOMO {d['tomo']} - Questão {d['numero']} | "
                f"Marcada: {d['alternativa_marcada']} | Correta: {d['alternativa_correta']}"
            )

        txt_content = "\n".join(linhas_txt)

        # -----------------------------
        # Nomes de arquivos
        # -----------------------------
        nome_json = "relatorio_estudante.json"
        nome_txt = f"relatorio_{st.session_state.registro_academico}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

        # -----------------------------
        # Tentativa de salvar em disco
        # (útil localmente; no Cloud é efêmero)
        # -----------------------------
        caminho_json = os.path.join(DIRETORIO_RELATORIOS, nome_json)
        caminho_txt = os.path.join(DIRETORIO_RELATORIOS, nome_txt)

        try:
            os.makedirs(DIRETORIO_RELATORIOS, exist_ok=True)

            with open(caminho_json, "w", encoding="utf-8") as f_json:
                f_json.write(json_content)

            with open(caminho_txt, "w", encoding="utf-8") as f_txt:
                f_txt.write(txt_content)
        except Exception as e_local:
            # Não falha o relatório por causa do disco no Streamlit Cloud
            caminho_json = None
            caminho_txt = None
            st.warning(f"Relatório gerado em memória, mas não foi possível salvar em disco: {e_local}")

        return True, {
            "json": caminho_json,
            "txt": caminho_txt,
            "nome_json": nome_json,
            "nome_txt": nome_txt,
            "json_content": json_content,
            "txt_content": txt_content,
            "aproveitamento": aproveitamento,
            "total_questoes": total_questoes,
            "total_respondidas": total_respondidas
        }

    except Exception as e:
        st.error(f"Erro ao gerar relatório: {e}")
        return False, None


# ============================
# INTERFACE
# ============================
def header():
    atualizar_tempo_prova()
   # st.caption(
    #    f"DEBUG modo_atual={st.session_state.get('modo_atual')} | timer_ativo={st.session_state.get('timer_ativo')} | tempo_restante={st.session_state.get('tempo_restante')}")

    st.title("📚 Aplicativo para o Processo Ensino / Aprendizagem")

    if st.session_state.cadastro_confirmado:
        st.markdown(
            f"**👤 {st.session_state.nome_usuario}** | "
            f"**📋 RA:** {st.session_state.registro_academico} | "
            f"**🏫 Turma:** {st.session_state.turma}"
        )

    col1, col2, col3 = st.columns([2, 2, 1])
    total = st.session_state.total_acertos + st.session_state.total_erros
    with col1:
        st.info(f"📊 Acertos: {st.session_state.total_acertos} | Erros: {st.session_state.total_erros} | Total: {total}")

    with col2:
        if st.session_state.get("modo_atual") == "prova":
            tempo_raw = st.session_state.get("tempo_restante", None)
            if tempo_raw is None:
                tempo_raw = st.session_state.get("tempo_prova", 0)

            try:
                tempo_seg = int(float(tempo_raw))
            except Exception:
                tempo_seg = 0

            tempo_seg = max(0, tempo_seg)
            mins, secs = divmod(tempo_seg, 60)

            #st.caption(
            #    f"DEBUG raw={tempo_raw} | tipo_raw={type(tempo_raw)} | "
            #    f"tempo_seg={tempo_seg} | mins={mins} | secs={secs}"
            #)

            # ✅ força render texto puro (sem st.warning)
            cor = "#ff4b4b" if tempo_seg <= 300 else "#f5c542"
            st.markdown(
                f"""
                <div style="
                    background: rgba(245,197,66,0.15);
                    border-radius: 10px;
                    padding: 14px 16px;
                    font-weight: 600;
                    font-size: 18px;
                    color: {cor};
                ">
                    ⏱ Tempo: {mins:02d}:{secs:02d}
                </div>
                """,
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                """
                <div style="
                    background: rgba(245,197,66,0.15);
                    border-radius: 10px;
                    padding: 14px 16px;
                    font-weight: 600;
                    font-size: 18px;
                    color: #f5c542;
                ">
                    ⏱ Tempo: 00:00
                </div>
                """,
                unsafe_allow_html=True
            )

    with col3:
        if st.button("💾 Salvar progresso"):
            salvar_progresso()

    if st.session_state.status_msg:
        st.caption(st.session_state.status_msg)


def tela_cadastro():
    st.subheader("🎓 Cadastro do Estudante")
    st.write("Preencha seus dados para iniciar.")

    with st.form("form_cadastro"):
        nome = st.text_input("Nome Completo", value=st.session_state.nome_usuario)
        ra = st.text_input("Registro Acadêmico (RA)", value=st.session_state.registro_academico)
        turma = st.text_input("Turma", value=st.session_state.turma)

        col1, col2 = st.columns(2)
        with col1:
            iniciar = st.form_submit_button("🚀 Iniciar Avaliação", width="stretch")
        with col2:
            carregar = st.form_submit_button("📂 Carregar dados salvos", width="stretch")

    if carregar:
        carregar_progresso()
        st.rerun()

    if iniciar:
        if not nome.strip():
            st.warning("Por favor, preencha o nome completo.")
            return
        if not ra.strip():
            st.warning("Por favor, preencha o Registro Acadêmico.")
            return
        if not turma.strip():
            st.warning("Por favor, preencha a turma.")
            return

        st.session_state.nome_usuario = nome.strip()
        st.session_state.registro_academico = ra.strip()
        st.session_state.turma = turma.strip()
        st.session_state.cadastro_confirmado = True
        salvar_progresso()
        st.rerun()


def painel_menu():
    st.subheader("📁 Carregar planilha e configurar estudo")

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        modo = st.radio(
            "Modo",
            options=["estudo", "prova"],
            index=0 if st.session_state.modo_atual == "estudo" else 1,
            horizontal=True
        )
    with col_m2:
        if modo == "prova":
            minutos = st.number_input(
                "Tempo da prova (minutos)",
                min_value=1,
                max_value=600,
                value=max(1, st.session_state.tempo_prova // 60),
                step=1
            )
            st.session_state.tempo_prova = int(minutos) * 60
            if not st.session_state.timer_ativo:
                st.session_state.tempo_restante = st.session_state.tempo_prova
        else:
            st.session_state.tempo_restante = None
            st.session_state.timer_ativo = False
            st.session_state.prova_inicio_epoch = None
            st.session_state.prova_fim_epoch = None
            st.session_state.questoes_iniciadas = False

    st.session_state.modo_atual = modo

    st.markdown("---")
    col_dir1, col_dir2 = st.columns(2)
    with col_dir1:
        st.session_state.diretorio_teorias = st.text_input(
            "📚 Diretório de Teorias (.docx)",
            value=st.session_state.diretorio_teorias
        )
    with col_dir2:
        st.session_state.diretorio_imagens = st.text_input(
            "🖼️ Diretório de Imagens",
            value=st.session_state.diretorio_imagens
        )

    st.markdown("---")
    st.write("### Planilha de questões")

    caminho_padrao = os.path.join(os.getcwd(), ARQUIVO_PLANILHA_PADRAO)
    st.caption(f"Arquivo padrão: {caminho_padrao}")

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("🚀 Carregar Planilha", width="stretch"):
            ok = carregar_planilha_arquivo(caminho_local=ARQUIVO_PLANILHA_PADRAO)
            if ok:
                st.rerun()

    with col2:
        if st.button("📊 Gerar Relatório", width="stretch"):
            ok, info = gerar_relatorio()
            if ok:
                st.success(
                    f"Relatório gerado com sucesso!\n\n"
                    f"- JSON: {info['json']}\n"
                    f"- TXT: {info['txt']}"
                )

    with col3:
        if st.button("🔄 Carregar Progresso", width="stretch"):
            carregar_progresso()
            st.rerun()

    if st.session_state.planilha_carregada and st.session_state.questoes_agrupadas:
        st.success(f"{len(st.session_state.questoes_agrupadas)} questões prontas para estudo.")
        if st.button("▶️ Ir para questões", width="stretch"):
            st.rerun()


def iniciar_timer_se_necessario():
    if st.session_state.modo_atual != "prova":
        return

    # Já ativo e com fim definido -> não reinicia
    if st.session_state.timer_ativo and st.session_state.prova_fim_epoch is not None:
        return

    agora = time.time()

    # Se já houver tempo_restante salvo, usa ele; senão usa tempo total da prova
    if st.session_state.tempo_restante is not None:
        restante = int(st.session_state.tempo_restante)
    else:
        restante = int(st.session_state.tempo_prova)

    st.session_state.timer_ativo = True
    st.session_state.prova_inicio_epoch = agora
    st.session_state.prova_fim_epoch = agora + restante
    st.session_state.tempo_restante = restante


def tela_questoes():
    if not st.session_state.questoes_agrupadas:
        st.warning("Nenhuma questão carregada.")
        return

    # ✅ Inicia timer quando entra na tela (modo prova)
    iniciar_timer_se_necessario()

    # ✅ Atualiza tempo a cada rerun
    atualizar_tempo_prova()

    # ✅ Auto refresh contínuo
    if st.session_state.modo_atual == "prova" and st.session_state.timer_ativo:
        st_autorefresh(interval=20000, limit=None, key="timer_autorefresh")

    if st.session_state.modo_atual == "prova" and st.session_state.tempo_restante == 0:
        st.warning("⏰ Tempo esgotado! Você pode continuar respondendo sem limite, se desejar.")

    total_q = len(st.session_state.questoes_agrupadas)
    idx = min(max(0, st.session_state.questao_atual), total_q - 1)
    st.session_state.questao_atual = idx
    questao = st.session_state.questoes_agrupadas[idx]

    st.markdown("---")
    st.subheader(f"📖 TOMO {questao['tomo']} - Questão {to_int_safe(questao['numero'])}")
    st.caption(f"Questão {idx + 1} de {total_q}")

    #if questao["usa_algarismos_romanos"]:
      #  st.info("ℹ️ Esta questão usa algarismos romanos (I, II, III...)")

    with st.expander("📄 Ver Enunciado Completo", expanded=True):
        st.write(questao["enunciado"] if questao["enunciado"] else "Enunciado não disponível.")
        if questao.get("imagem"):
            mostrar_imagem_questao_streamlit(questao["imagem"])

    # ======= Alternativas (BLOQUEIO REAL: some o radio após verificar) =======
    st.markdown("### Escolha uma alternativa")
    opcoes = [alt["letra"] for alt in questao["alternativas"]]
    mapa_texto = {alt["letra"]: f"{alt['letra']}) {alt['texto']}" for alt in questao["alternativas"]}

    ja_verificada = idx in st.session_state.questoes_respondidas
    resposta_congelada = st.session_state.respostas_por_questao.get(str(idx), "")

    if ja_verificada:
        st.info("🔒 Questão já verificada. A resposta está bloqueada.")
        for alt in questao["alternativas"]:
            letra = alt["letra"]
            texto = f"{letra}) {alt['texto']}"
            if letra == resposta_congelada:
                st.warning(f" {texto}  *(resposta registrada)*")
            else:
                st.write(texto)
        alt_sel = resposta_congelada
    else:
        valor_inicial = st.session_state.respostas_por_questao.get(str(idx), "")
        key_radio = f"resp_q_{idx}_v{st.session_state.radio_reset_version}"

        if key_radio not in st.session_state:
            st.session_state[key_radio] = valor_inicial if valor_inicial in opcoes else ""

        opcoes_com_vazio = [""] + opcoes

        def radio_format(x):
            return "Selecione..." if x == "" else mapa_texto.get(x, x)

        st.radio(
            "Alternativas",
            options=opcoes_com_vazio,
            key=key_radio,
            format_func=radio_format,
            label_visibility="collapsed"
        )
        alt_sel = st.session_state.get(key_radio, "")

    # ======= Botões =======
    col1, col2, col3, col4, col5 = st.columns([1, 1, 1, 1, 1])

    with col1:
        if st.button("⬅ Anterior", width="stretch", disabled=(idx == 0)):
            st.session_state.questao_atual = max(0, idx - 1)
            st.session_state.resposta_verificada = False
            st.session_state.mostrar_justificativas = False
            st.rerun()

    with col2:
        if st.button("✅ Verificar", width="stretch", disabled=ja_verificada):
            verificar_resposta_streamlit()
            st.rerun()

    with col3:
        pode_mostrar_resposta = st.session_state.resposta_verificada or ja_verificada
        if st.button("🔍 Mostrar Resposta", width="stretch", disabled=not pode_mostrar_resposta):
            st.session_state.mostrar_justificativas = True
            st.rerun()

    with col4:
        if st.button("📖 Ver Teoria", width="stretch"):
            st.session_state["abrir_teoria_para_idx"] = idx
            st.rerun()

    with col5:
        if st.button("➡ Próxima", width="stretch", disabled=not ja_verificada):
            if idx < total_q - 1:
                st.session_state.questao_atual = idx + 1
                st.session_state.resposta_verificada = False
                st.session_state.mostrar_justificativas = False
                st.rerun()
            else:
                st.info("Você chegou ao fim das questões.")

    # ======= Feedback =======
    if ja_verificada and resposta_congelada:
        renderizar_feedback(questao, resposta_congelada, idx=idx)
    else:
        if st.session_state.resposta_verificada and alt_sel:
            renderizar_feedback(questao, alt_sel, idx=idx)

    # Justificativas
    if st.session_state.mostrar_justificativas:
        renderizar_justificativas(questao)

    # Teoria
    if st.session_state.get("abrir_teoria_para_idx", None) == idx:
        with st.expander("📖 Teoria da questão (expandida)", expanded=True):
            renderizar_teoria(questao)

    st.markdown("---")

    c1, c2, c3, c4 = st.columns([1, 1, 1, 1])

    with c1:
        if st.button("🏠 Menu Principal", width="stretch"):
            st.session_state.planilha_carregada = False
            st.session_state.df = None
            st.session_state.questoes_agrupadas = []
            st.session_state.questao_atual = 0
            st.session_state.resposta_verificada = False
            st.session_state.alternativa_selecionada = None
            st.session_state.mostrar_justificativas = False
            st.session_state.abrir_teoria_para_idx = None
            st.session_state.questoes_iniciadas = False
            st.session_state.prova_inicio_epoch = None
            st.session_state.prova_fim_epoch = None
            st.session_state.timer_ativo = False
            st.session_state.tempo_restante = None if st.session_state.modo_atual == "estudo" else st.session_state.tempo_restante
            st.rerun()

    with c2:
        if st.button("📊 Gerar Relatório", width="stretch"):
            ok, info = gerar_relatorio()
            if ok:
                st.success(
                    f"✅ Relatório gerado!\n"
                    f"Aproveitamento: {info['aproveitamento']:.1f}%\n"
                    f"Arquivos: {info['json']} | {info['txt']}"
                )

    with c3:
        todas_respondidas = len(st.session_state.questoes_respondidas) == len(st.session_state.questoes_agrupadas)
        if st.button("🏁 Finalizar Avaliação", width="stretch", disabled=not todas_respondidas):
            ok, info = gerar_relatorio()
            if ok:
                st.balloons()
                st.success(
                    f"🎉 Avaliação finalizada!\n\n"
                    f"Total: {info['total_questoes']}\n"
                    f"Respondidas: {info['total_respondidas']}\n"
                    f"Acertos: {st.session_state.total_acertos}\n"
                    f"Erros: {st.session_state.total_erros}\n"
                    f"Aproveitamento: {info['aproveitamento']:.1f}%"
                )
    with c4:
        if st.button("📧 Enviar Relatório", width="stretch"):
            ok, info = gerar_relatorio()
            if ok:
                ok_email, msg_email = enviar_relatorio_por_email(
                    txt_content=info["txt_content"],
                    json_content=info["json_content"],
                    nome_txt=info["nome_txt"],
                    nome_json=info["nome_json"],
                )
                if ok_email:
                    st.success(msg_email)
                else:
                    st.error(msg_email)

                # ✅ COLOQUE O TRECHO DE DOWNLOAD AQUI
                st.download_button(
                    "⬇️ Baixar TXT",
                    data=info["txt_content"].encode("utf-8"),
                    file_name=info["nome_txt"],
                    mime="text/plain",
                    width="stretch"
                )
                st.download_button(
                    "⬇️ Baixar JSON",
                    data=info["json_content"].encode("utf-8"),
                    file_name=info["nome_json"],
                    mime="application/json",
                    width="stretch"
                )


# ============================
# MAIN
# ============================
def main():
    init_session()

    # ✅ Atualiza tempo antes do header para o topo mostrar valor correto
    atualizar_tempo_prova()

    header()

    if not st.session_state.cadastro_confirmado:
        tela_cadastro()
        return

    if not st.session_state.planilha_carregada:
        painel_menu()
        return

    tela_questoes()
    
if __name__ == "__main__":
    main()


