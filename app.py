import streamlit as st
import tempfile
import io
import contextlib
import zipfile
from pathlib import Path

from main import processar_contas, processar_extratos


# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Automação Condomínio",
    page_icon="🏢",
    layout="wide",
)


# ── ESTADO DA SESSÃO ──────────────────────────────────────────────────────────
def _inicializar_estado():
    if "tempdir" not in st.session_state:
        st.session_state.tempdir = tempfile.TemporaryDirectory()
    if "resultado_contas" not in st.session_state:
        st.session_state.resultado_contas = None
    if "resultado_extratos" not in st.session_state:
        st.session_state.resultado_extratos = None
    if "processado" not in st.session_state:
        st.session_state.processado = False


_inicializar_estado()
base_dir = Path(st.session_state.tempdir.name)


# ── HELPERS ───────────────────────────────────────────────────────────────────
def _salvar_uploads(arquivos, pasta: Path):
    pasta.mkdir(parents=True, exist_ok=True)
    for arq in arquivos:
        (pasta / arq.name).write_bytes(arq.getbuffer())


def _construir_zip(resultado_contas: dict, resultado_extratos: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for pdf in resultado_contas.get("arquivos_gerados", []):
            p = Path(pdf)
            if p.exists():
                zf.write(p, f"Agrupado/{p.name}")

        for chave in ("arquivo_excel", "arquivo_html_corrente", "arquivo_html_poupanca"):
            arq = resultado_extratos.get(chave)
            if arq:
                p = Path(arq)
                if p.exists():
                    zf.write(p, f"Extratos/{p.name}")

    buf.seek(0)
    return buf.read()


# ── INTERFACE — UPLOAD ────────────────────────────────────────────────────────
st.title("Automação Condomínio")
st.caption("Processamento de contas, comprovantes e extratos bancários")

aba_contas, aba_extratos = st.tabs(["Contas & Comprovantes", "Extratos Bancários"])

with aba_contas:
    st.subheader("Upload de PDFs")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Contas** (PDFs)")
        st.caption("Nomeie como `nome_conta.pdf`")
        uploads_contas = st.file_uploader(
            "Contas",
            type="pdf",
            accept_multiple_files=True,
            key="upload_contas",
            label_visibility="collapsed",
        )

    with col2:
        st.markdown("**Comprovantes** (PDFs, opcional)")
        st.caption("Nomeie como `nome_comprovante.pdf`")
        uploads_comprovantes = st.file_uploader(
            "Comprovantes",
            type="pdf",
            accept_multiple_files=True,
            key="upload_comprovantes",
            label_visibility="collapsed",
        )

with aba_extratos:
    st.subheader("Upload dos Extratos HTML")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Extrato Conta Corrente**")
        upload_corrente = st.file_uploader(
            "Extrato Corrente",
            type=["html", "htm"],
            key="upload_corrente",
            label_visibility="collapsed",
        )

    with col2:
        st.markdown("**Extrato Conta Poupança**")
        upload_poupanca = st.file_uploader(
            "Extrato Poupança",
            type=["html", "htm"],
            key="upload_poupanca",
            label_visibility="collapsed",
        )


# ── VALIDAÇÃO ─────────────────────────────────────────────────────────────────
uploads_contas = uploads_contas or []
uploads_comprovantes = uploads_comprovantes or []

tem_contas = len(uploads_contas) > 0
tem_corrente = upload_corrente is not None
tem_poupanca = upload_poupanca is not None
pode_processar = tem_contas and tem_corrente and tem_poupanca

if not pode_processar:
    faltando = []
    if not tem_contas:
        faltando.append("pelo menos 1 PDF de conta")
    if not tem_corrente:
        faltando.append("extrato corrente (HTML)")
    if not tem_poupanca:
        faltando.append("extrato poupança (HTML)")
    st.info("Para habilitar o processamento, envie: " + " · ".join(faltando))


# ── EXECUÇÃO ──────────────────────────────────────────────────────────────────
st.divider()

if st.button(
    "Processar",
    disabled=not pode_processar,
    type="primary",
    use_container_width=True,
):
    # Reseta resultados anteriores
    st.session_state.processado = False
    st.session_state.resultado_contas = None
    st.session_state.resultado_extratos = None

    # Salva uploads nas pastas esperadas pelo main.py
    _salvar_uploads(uploads_contas, base_dir / "Contas")
    _salvar_uploads(uploads_comprovantes, base_dir / "Comprovantes")

    pasta_entrada = base_dir / "Extratos" / "Entrada"
    pasta_entrada.mkdir(parents=True, exist_ok=True)
    (pasta_entrada / "Extrato Corrente.html").write_bytes(upload_corrente.getbuffer())
    (pasta_entrada / "Extrato Poupança.html").write_bytes(upload_poupanca.getbuffer())

    erro_etapa1 = None

    # Etapa 1 — Agrupamento
    with st.status("Etapa 1 — Agrupando contas e comprovantes...", expanded=True) as s1:
        try:
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                resultado_contas = processar_contas(base_dir)
            st.code(buf.getvalue(), language="text")
            s1.update(label="Etapa 1 concluída!", state="complete")
            st.session_state.resultado_contas = resultado_contas
        except Exception as exc:
            s1.update(label=f"Etapa 1 — Erro: {exc}", state="error")
            st.error(f"Erro na etapa 1: {exc}")
            erro_etapa1 = exc

    # Etapa 2 — Conferência de extratos (só executa se etapa 1 passou)
    if erro_etapa1 is None:
        with st.status("Etapa 2 — Conferindo extratos bancários...", expanded=True) as s2:
            try:
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf):
                    resultado_extratos = processar_extratos(base_dir)
                st.code(buf.getvalue(), language="text")
                s2.update(label="Etapa 2 concluída!", state="complete")
                st.session_state.resultado_extratos = resultado_extratos
                st.session_state.processado = True
            except Exception as exc:
                s2.update(label=f"Etapa 2 — Erro: {exc}", state="error")
                st.error(f"Erro na etapa 2: {exc}")


# ── RESULTADOS ────────────────────────────────────────────────────────────────
if st.session_state.processado:
    res_c = st.session_state.resultado_contas
    res_e = st.session_state.resultado_extratos

    st.divider()
    st.subheader("Resultados")

    # Métricas
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Contas processadas", res_c["total_processadas"])
    col2.metric("Erros (contas)", res_c["total_erros"])
    col3.metric("Match forte", res_e["match_forte"])
    col4.metric("Match ambíguo", res_e["match_ambiguo"])
    col5.metric("Sem match", res_e["sem_match"])

    st.divider()

    # Downloads individuais
    col_pdfs, col_outros = st.columns(2)

    with col_pdfs:
        st.markdown("**PDFs agrupados**")
        pdfs = res_c.get("arquivos_gerados", [])
        if pdfs:
            for pdf_path in pdfs:
                p = Path(pdf_path)
                if p.exists():
                    st.download_button(
                        label=f"Baixar {p.name}",
                        data=p.read_bytes(),
                        file_name=p.name,
                        mime="application/pdf",
                        key=f"dl_pdf_{p.name}",
                    )
        else:
            st.caption("Nenhum PDF gerado.")

    with col_outros:
        st.markdown("**Conferência e extratos destacados**")

        excel_path = Path(res_e.get("arquivo_excel", ""))
        if excel_path.exists():
            st.download_button(
                label=f"Baixar {excel_path.name}",
                data=excel_path.read_bytes(),
                file_name=excel_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_excel",
            )

        html_corrente = Path(res_e.get("arquivo_html_corrente", ""))
        if html_corrente.exists():
            st.download_button(
                label=f"Baixar {html_corrente.name}",
                data=html_corrente.read_bytes(),
                file_name=html_corrente.name,
                mime="text/html",
                key="dl_html_corrente",
            )

        html_poupanca = Path(res_e.get("arquivo_html_poupanca", ""))
        if html_poupanca.exists():
            st.download_button(
                label=f"Baixar {html_poupanca.name}",
                data=html_poupanca.read_bytes(),
                file_name=html_poupanca.name,
                mime="text/html",
                key="dl_html_poupanca",
            )

    st.divider()

    # Download ZIP
    st.markdown("**Baixar tudo de uma vez**")
    zip_bytes = _construir_zip(res_c, res_e)
    st.download_button(
        label="Baixar tudo em ZIP",
        data=zip_bytes,
        file_name="resultado_condominio.zip",
        mime="application/zip",
        type="primary",
        use_container_width=True,
        key="dl_zip",
    )
