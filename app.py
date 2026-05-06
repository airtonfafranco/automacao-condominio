import streamlit as st
import tempfile
import io
import json
import contextlib
import zipfile
from pathlib import Path

from main import processar_contas, processar_extratos
from gerar_balancete import gerar_balancete, MESES_PT


# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Automação Condomínio",
    page_icon="🏢",
    layout="wide",
)


# ── ESTADO DA SESSÃO ──────────────────────────────────────────────────────────
def _inicializar_estado():
    defaults = {
        "tempdir":                None,
        "resultado_contas":       None,
        "resultado_extratos":     None,
        "processado":             False,
        "resultado_balancete":    None,
        "ambiguidade_candidatos": None,
        "ambiguidade_shortfall":  None,
        "balancete_processado":   False,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val
    if st.session_state.tempdir is None:
        st.session_state.tempdir = tempfile.TemporaryDirectory()


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


def _num_mes(nome: str) -> int:
    for num, nm in MESES_PT.items():
        if nm == nome:
            return num
    return 1


# ── TÍTULO ────────────────────────────────────────────────────────────────────
st.title("Automação Condomínio")
st.caption("Processamento de contas, comprovantes, extratos bancários e balancete")

aba_contas, aba_extratos, aba_balancete = st.tabs(
    ["Contas & Comprovantes", "Extratos Bancários", "Balancete"]
)


# ══════════════════════════════════════════════════════════════════════════════
# ABA 1 — CONTAS & COMPROVANTES
# ══════════════════════════════════════════════════════════════════════════════
with aba_contas:
    st.subheader("Upload de PDFs")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Contas** (PDFs)")
        st.caption("Nomeie como `nome_conta.pdf`")
        uploads_contas = st.file_uploader(
            "Contas", type="pdf", accept_multiple_files=True,
            key="upload_contas", label_visibility="collapsed",
        )

    with col2:
        st.markdown("**Comprovantes** (PDFs, opcional)")
        st.caption("Nomeie como `nome_comprovante.pdf`")
        uploads_comprovantes = st.file_uploader(
            "Comprovantes", type="pdf", accept_multiple_files=True,
            key="upload_comprovantes", label_visibility="collapsed",
        )


# ══════════════════════════════════════════════════════════════════════════════
# ABA 2 — EXTRATOS BANCÁRIOS
# ══════════════════════════════════════════════════════════════════════════════
with aba_extratos:
    st.subheader("Upload dos Extratos HTML")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Extrato Conta Corrente**")
        upload_corrente = st.file_uploader(
            "Extrato Corrente", type=["html", "htm"],
            key="upload_corrente", label_visibility="collapsed",
        )

    with col2:
        st.markdown("**Extrato Conta Poupança**")
        upload_poupanca = st.file_uploader(
            "Extrato Poupança", type=["html", "htm"],
            key="upload_poupanca", label_visibility="collapsed",
        )


# ══════════════════════════════════════════════════════════════════════════════
# ABA 3 — BALANCETE
# ══════════════════════════════════════════════════════════════════════════════
with aba_balancete:
    st.subheader("Parâmetros do Balancete")

    col1, col2, col3 = st.columns(3)
    with col1:
        mes_nome_sel = st.selectbox("Mês de referência", list(MESES_PT.values()), index=2)
        mes_num_sel  = _num_mes(mes_nome_sel)
    with col2:
        ano_sel = st.number_input("Ano", min_value=2020, max_value=2035,
                                  value=2026, step=1)
    with col3:
        caixa_fin_sel = st.number_input("Caixa Final (R$)",    value=0.00,
                                        format="%.2f", step=0.01)
        caixa_ant_sel = st.number_input("Caixa Anterior (R$)", value=0.00,
                                        format="%.2f", step=0.01)

    st.divider()

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**Planilha TACON/TUV** (.xlsx) — obrigatório")
        upload_tacon = st.file_uploader(
            "TACON_TUV", type=["xlsx"],
            key="upload_tacon", label_visibility="collapsed",
        )
    with col_b:
        st.markdown("**Inadimplentes mês anterior** (.json) — opcional")
        st.caption("Arquivo gerado automaticamente no balancete do mês anterior.")
        upload_inad_ant = st.file_uploader(
            "Inadimplentes anterior", type=["json"],
            key="upload_inad_ant", label_visibility="collapsed",
        )

    # ── Resolução de ambiguidade ──────────────────────────────────────────────
    if st.session_state.ambiguidade_candidatos:
        st.divider()
        st.warning(
            f"⚠️ **Inadimplência ambígua** — shortfall de "
            f"R$ {st.session_state.ambiguidade_shortfall:.2f} corresponde a "
            f"{len(st.session_state.ambiguidade_candidatos)} unidades possíveis. "
            "Selecione a unidade que **não pagou**:"
        )

        candidatos = st.session_state.ambiguidade_candidatos
        col_unit = next(
            (c for c in candidatos[0].keys() if c.strip().upper() == "UNIDADE"),
            list(candidatos[0].keys())[0],
        )

        def _label(c, i):
            t = c.get(next((k for k in c if k.upper() == "TACON"), ""), 0) or 0
            v = c.get(next((k for k in c if k.upper() == "TUV"),   ""), 0) or 0
            return (f"{c.get(col_unit, f'Unidade {i+1}')}  "
                    f"(TACON R$ {float(t):.2f} + TUV R$ {float(v):.2f})")

        opcoes = {_label(c, i): c for i, c in enumerate(candidatos)}
        sel_label = st.selectbox("Unidade inadimplente:", list(opcoes.keys()),
                                 key="sel_inadimplente")

        if st.button("✅ Confirmar e Gerar Balancete", type="primary",
                     key="btn_confirmar_inad"):
            escolhido = dict(opcoes[sel_label])
            escolhido["unidade"] = escolhido.get(col_unit, "?")

            pasta_entrada = base_dir / "Extratos" / "Entrada"
            arq_tacon_path = pasta_entrada / "_tacon_tuv.xlsx"

            inad_ant_dados = None
            if upload_inad_ant:
                inad_ant_dados = json.loads(upload_inad_ant.getvalue().decode("utf-8"))

            with st.status("Gerando balancete com inadimplente confirmado...",
                           expanded=True) as s3:
                try:
                    buf = io.StringIO()
                    with contextlib.redirect_stdout(buf):
                        resultado = gerar_balancete(
                            base_dir             = base_dir,
                            mes_nome             = mes_nome_sel,
                            mes_num              = mes_num_sel,
                            ano                  = int(ano_sel),
                            caixa_fin            = float(caixa_fin_sel),
                            caixa_ant            = float(caixa_ant_sel),
                            arq_tacon_tuv        = arq_tacon_path,
                            inad_ant_dados       = inad_ant_dados,
                            inadimplente_forcado = escolhido,
                        )
                    st.code(buf.getvalue(), language="text")
                    s3.update(label="Balancete gerado com sucesso!", state="complete")
                    st.session_state.resultado_balancete    = resultado
                    st.session_state.ambiguidade_candidatos = None
                    st.session_state.ambiguidade_shortfall  = None
                    st.session_state.balancete_processado   = True
                except Exception as exc:
                    s3.update(label=f"Erro: {exc}", state="error")
                    st.error(f"Erro ao gerar balancete: {exc}")

    st.divider()

    # ── Botão principal ───────────────────────────────────────────────────────
    pode_balancete = st.session_state.processado and upload_tacon is not None

    if not st.session_state.processado:
        st.info("Execute primeiro as **Etapas 1 e 2** (botão 'Processar' abaixo).")
    elif upload_tacon is None:
        st.info("Faça o upload da planilha **TACON/TUV** para habilitar o balancete.")

    if st.button(
        "Gerar Balancete",
        disabled=not pode_balancete,
        type="primary",
        use_container_width=True,
        key="btn_balancete",
    ):
        # Resetar estado anterior
        st.session_state.resultado_balancete    = None
        st.session_state.ambiguidade_candidatos = None
        st.session_state.ambiguidade_shortfall  = None
        st.session_state.balancete_processado   = False

        # Salvar TACON_TUV no tempdir
        pasta_entrada = base_dir / "Extratos" / "Entrada"
        pasta_entrada.mkdir(parents=True, exist_ok=True)
        arq_tacon_path = pasta_entrada / "_tacon_tuv.xlsx"
        arq_tacon_path.write_bytes(upload_tacon.getbuffer())

        # Inadimplentes do mês anterior (JSON, opcional)
        inad_ant_dados = None
        if upload_inad_ant:
            inad_ant_dados = json.loads(upload_inad_ant.getvalue().decode("utf-8"))

        with st.status("Etapa 3 — Gerando balancete...", expanded=True) as s3:
            try:
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf):
                    resultado = gerar_balancete(
                        base_dir      = base_dir,
                        mes_nome      = mes_nome_sel,
                        mes_num       = mes_num_sel,
                        ano           = int(ano_sel),
                        caixa_fin     = float(caixa_fin_sel),
                        caixa_ant     = float(caixa_ant_sel),
                        arq_tacon_tuv = arq_tacon_path,
                        inad_ant_dados = inad_ant_dados,
                    )
                st.code(buf.getvalue(), language="text")

                if resultado["ambiguidade"]:
                    s3.update(
                        label="⚠️ Inadimplência ambígua — confirme a unidade acima.",
                        state="error",
                    )
                    st.session_state.ambiguidade_candidatos = resultado["candidatos"]
                    st.session_state.ambiguidade_shortfall  = resultado["shortfall"]
                    st.rerun()
                else:
                    s3.update(label="Balancete gerado com sucesso!", state="complete")
                    st.session_state.resultado_balancete  = resultado
                    st.session_state.balancete_processado = True

            except Exception as exc:
                s3.update(label=f"Erro: {exc}", state="error")
                st.error(f"Erro ao gerar balancete: {exc}")

    # ── Resultados do balancete ───────────────────────────────────────────────
    if st.session_state.balancete_processado:
        res    = st.session_state.resultado_balancete
        resumo = res.get("resumo", {})

        st.divider()
        st.subheader("Resumo do Balancete")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Despesas",   f"R$ {resumo.get('total_despesas', 0):,.2f}")
        col2.metric("Total Receitas",   f"R$ {resumo.get('total_receitas', 0):,.2f}")
        col3.metric("Saldo do Mês",     f"R$ {resumo.get('saldo_mes', 0):,.2f}")
        col4.metric("Saldo Disponível", f"R$ {resumo.get('saldo_disponivel', 0):,.2f}")

        st.divider()
        col_esq, col_dir = st.columns(2)

        with col_esq:
            st.markdown("**Detalhamento de Receitas**")
            st.write(f"TACON: R$ {resumo.get('tacon', 0):,.2f}")
            st.write(f"TUV: R$ {resumo.get('tuv', 0):,.2f}")
            if resumo.get("multa", 0) > 0:
                st.write(f"Multa (atraso): R$ {resumo.get('multa', 0):,.2f}")
            if resumo.get("ref_anterior", 0) > 0:
                st.write(f"Ref. mês anterior: R$ {resumo.get('ref_anterior', 0):,.2f}")
            st.write(f"Juros Poupança: R$ {resumo.get('juros_poupanca', 0):,.2f}")

        with col_dir:
            st.markdown("**Downloads**")

            xlsx_path = Path(res.get("arquivo_gerado") or "")
            if xlsx_path.exists():
                st.download_button(
                    label=f"📥 Baixar {xlsx_path.name}",
                    data=xlsx_path.read_bytes(),
                    file_name=xlsx_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_balancete_xlsx",
                )

            inad_path = Path(res.get("inadimplentes_arquivo") or "")
            if inad_path.exists():
                st.download_button(
                    label=f"📥 Baixar {inad_path.name}",
                    data=inad_path.read_bytes(),
                    file_name=inad_path.name,
                    mime="application/json",
                    key="dl_inad_json",
                )
                st.caption(
                    "⬆️ Guarde este arquivo JSON. "
                    "Ele será o input de **Inadimplentes mês anterior** no próximo mês."
                )


# ══════════════════════════════════════════════════════════════════════════════
# VALIDAÇÃO ETAPAS 1 + 2
# ══════════════════════════════════════════════════════════════════════════════
uploads_contas       = uploads_contas or []
uploads_comprovantes = uploads_comprovantes or []

tem_contas     = len(uploads_contas) > 0
tem_corrente   = upload_corrente is not None
tem_poupanca   = upload_poupanca is not None
pode_processar = tem_contas and tem_corrente and tem_poupanca

if not pode_processar:
    faltando = []
    if not tem_contas:   faltando.append("pelo menos 1 PDF de conta")
    if not tem_corrente: faltando.append("extrato corrente (HTML)")
    if not tem_poupanca: faltando.append("extrato poupança (HTML)")
    st.info("Para habilitar o processamento, envie: " + " · ".join(faltando))


# ══════════════════════════════════════════════════════════════════════════════
# BOTÃO PROCESSAR (Etapas 1 + 2)
# ══════════════════════════════════════════════════════════════════════════════
st.divider()

if st.button(
    "Processar",
    disabled=not pode_processar,
    type="primary",
    use_container_width=True,
):
    # Resetar tudo
    st.session_state.processado             = False
    st.session_state.resultado_contas       = None
    st.session_state.resultado_extratos     = None
    st.session_state.balancete_processado   = False
    st.session_state.resultado_balancete    = None
    st.session_state.ambiguidade_candidatos = None

    _salvar_uploads(uploads_contas,       base_dir / "Contas")
    _salvar_uploads(uploads_comprovantes, base_dir / "Comprovantes")

    pasta_entrada = base_dir / "Extratos" / "Entrada"
    pasta_entrada.mkdir(parents=True, exist_ok=True)
    (pasta_entrada / "Extrato Corrente.html").write_bytes(upload_corrente.getbuffer())
    (pasta_entrada / "Extrato Poupança.html").write_bytes(upload_poupanca.getbuffer())

    erro_etapa1 = None

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


# ══════════════════════════════════════════════════════════════════════════════
# RESULTADOS — ETAPAS 1 + 2
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.processado:
    res_c = st.session_state.resultado_contas
    res_e = st.session_state.resultado_extratos

    st.divider()
    st.subheader("Resultados — Contas e Extratos")

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Contas processadas", res_c["total_processadas"])
    col2.metric("Erros (contas)",     res_c["total_erros"])
    col3.metric("Match forte",        res_e["match_forte"])
    col4.metric("Match ambíguo",      res_e["match_ambiguo"])
    col5.metric("Sem match",          res_e["sem_match"])

    st.divider()
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
    st.markdown("**Baixar tudo de uma vez (Etapas 1 + 2)**")
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

    if not st.session_state.balancete_processado:
        st.info(
            "✅ Etapas 1 e 2 concluídas. "
            "Vá para a aba **Balancete** para gerar o balancete do mês."
        )
