import re

with open("main.py", "r", encoding="utf-8") as f:
    content = f.read()

# ==============================================================
# TICKET 2: ORIGIN SELECTBOX
# ==============================================================
# The old origin block is from "if aba == "Nova Cotação":" to the end of info_o = buscar_endereco_cep(cep_origem)
old_origin_block = """    # Barra elegante e compacta de Origem
    col_orig_banner, col_orig_btn = st.columns([3.5, 1])
    with col_orig_banner:
        st.markdown(
            f'''
            <div style="background-color:#f8fafc;padding:6px 12px;border-radius:6px;border:1px solid #e2e8f0;display:flex;align-items:center;gap:12px;margin-bottom:16px;">
              <div style="background-color:#e0e7ff;color:#4f46e5;padding:6px;border-radius:4px;">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path><circle cx="12" cy="10" r="3"></circle></svg>
              </div>
              <div>
                <div style="font-weight:700;color:#0f172a;font-size:0.88rem;letter-spacing:-0.01em;">
                  ORIGEM: Matriz Londrina/PR - Next Cable
                </div>
                <div style="color:#64748b;font-size:0.77rem;margin-top:2px;">
                  CNPJ: 26.434.839/0001-21 | CEP: 86.071-000
                </div>
              </div>
            </div>
            ''',
            unsafe_allow_html=True
        )

    with col_orig_btn:
        with st.expander("Alterar Origem"):
            cnpj_remetente = st.text_input("CNPJ Origem", value=cnpj_rem_default, key="cnpj_rem_inp")
            cep_origem = st.text_input("CEP Origem", value=cep_rem_default, key="cep_o_inp")
            st.session_state.cnpj_rem_custom = cnpj_remetente
            st.session_state.cep_rem_custom = cep_origem

    cnpj_remetente = st.session_state.get("cnpj_rem_custom", "")
    cep_origem = st.session_state.get("cep_rem_custom", "")
    info_o = buscar_endereco_cep(cep_origem)"""


new_origin_block = """    # Barra elegante e compacta de Origem
    origens_dict = {
        "NEXT Matriz (Londrina/PR)": {"cnpj": "26434839000121", "cep": "86071000"},
        "R NET Telecom (Londrina/PR)": {"cnpj": "11275512000187", "cep": "86010070"},
        "GB Souza (Londrina/PR)": {"cnpj": "11572216000148", "cep": "86026090"},
        "NEXT Filial MG (Belo Horizonte/MG)": {"cnpj": "26434839000474", "cep": "30810600"},
        "NEXT Filial GO (Anápolis/GO)": {"cnpj": "26434839000202", "cep": "75114300"},
        "Outra Origem (Digitar Manualmente)": {"cnpj": "", "cep": ""}
    }

    col_orig_banner, col_orig_btn = st.columns([3.5, 1])
    with col_orig_btn:
        origem_selecionada = st.selectbox("Selecione a Origem", list(origens_dict.keys()), key="origem_select", label_visibility="collapsed")
    
    with col_orig_banner:
        if origem_selecionada == "Outra Origem (Digitar Manualmente)":
            c_cnpj, c_cep = st.columns(2)
            with c_cnpj:
                cnpj_remetente = st.text_input("CNPJ Origem", key="cnpj_rem_inp")
            with c_cep:
                cep_origem = st.text_input("CEP Origem", key="cep_o_inp")
        else:
            cnpj_remetente = origens_dict[origem_selecionada]["cnpj"]
            cep_origem = origens_dict[origem_selecionada]["cep"]
            st.markdown(
                f'''
                <div style="background-color:#f8fafc;padding:6px 12px;border-radius:6px;border:1px solid #e2e8f0;display:flex;align-items:center;gap:12px;margin-bottom:16px;">
                  <div style="background-color:#e0e7ff;color:#4f46e5;padding:6px;border-radius:4px;">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path><circle cx="12" cy="10" r="3"></circle></svg>
                  </div>
                  <div>
                    <div style="font-weight:700;color:#0f172a;font-size:0.88rem;letter-spacing:-0.01em;">
                      ORIGEM: {origem_selecionada}
                    </div>
                    <div style="color:#64748b;font-size:0.77rem;margin-top:2px;">
                      CNPJ: {cnpj_remetente} | CEP: {cep_origem}
                    </div>
                  </div>
                </div>
                ''',
                unsafe_allow_html=True
            )
            
    st.session_state.cnpj_rem_custom = cnpj_remetente
    st.session_state.cep_rem_custom = cep_origem
    info_o = buscar_endereco_cep(cep_origem)"""

if old_origin_block in content:
    content = content.replace(old_origin_block, new_origin_block)
else:
    print("WARNING: Could not find old origin block.")

# ==============================================================
# TICKET 1: VOLUME DATA EDITOR (Solução B)
# ==============================================================
old_volume_block = """    # 2. Grade de Volumes Dinâmica (Data Editor)
    with st.container(border=True):
        col_vol_tit, col_vol_act = st.columns([4, 1])
        with col_vol_tit:
            st.markdown("#### 2. Volumes da Cotação")
        with col_vol_act:
            if st.button("Adicionar Volume", use_container_width=True):
                if "volumes_data" not in st.session_state or not st.session_state.volumes_data:
                    st.session_state.volumes_data = [
                        {"Qtd": 1, "Alt (cm)": 30.0, "Larg (cm)": 30.0, "Comp (cm)": 40.0, "Peso Un. (kg)": 10.0}
                    ]
                else:
                    st.session_state.volumes_data.append(
                        {"Qtd": 1, "Alt (cm)": 30.0, "Larg (cm)": 30.0, "Comp (cm)": 40.0, "Peso Un. (kg)": 10.0}
                    )
                st.rerun()

        if "volumes_data" not in st.session_state:
            st.session_state.volumes_data = []

        if len(st.session_state.volumes_data) == 0:
            df_volumes_input = pd.DataFrame(columns=["Qtd", "Alt (cm)", "Larg (cm)", "Comp (cm)", "Peso Total (kg)"])
        else:
            df_volumes_input = pd.DataFrame(st.session_state.volumes_data)

        edited_volumes = st.data_editor(
            df_volumes_input,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="volumes_editor",
            column_config={
                "Qtd": st.column_config.NumberColumn("Qtd", min_value=1, step=1, default=1),
                "Alt (cm)": st.column_config.NumberColumn("Alt (cm)", min_value=1.0, step=5.0, default=30.0),
                "Larg (cm)": st.column_config.NumberColumn("Larg (cm)", min_value=1.0, step=5.0, default=30.0),
                "Comp (cm)": st.column_config.NumberColumn("Comp (cm)", min_value=1.0, step=5.0, default=40.0),
                "Peso Total (kg)": st.column_config.NumberColumn("Peso Total (kg)", min_value=0.1, step=0.5, default=10.0, format="%.2f"),
            },
        )

        if edited_volumes is not None:
            # Sync edited volumes to session state for resume
            st.session_state.volumes_data = edited_volumes.to_dict(orient="records")"""

new_volume_block = """    # 2. Grade de Volumes Dinâmica (Data Editor)
    with st.container(border=True):
        st.markdown("#### 2. Volumes da Cotação")

        # Configurar estado âncora do dataframe para evitar reset na tecla TAB
        if "volumes_df" not in st.session_state:
            if "volumes_data" in st.session_state and len(st.session_state.volumes_data) > 0:
                import pandas as pd
                st.session_state.volumes_df = pd.DataFrame(st.session_state.volumes_data)
            else:
                import pandas as pd
                st.session_state.volumes_df = pd.DataFrame(columns=["Qtd", "Alt (cm)", "Larg (cm)", "Comp (cm)", "Peso Total (kg)"])

        edited_volumes = st.data_editor(
            st.session_state.volumes_df,
            num_rows="dynamic",
            width="stretch",
            hide_index=True,
            key="volumes_editor",
            column_config={
                "Qtd": st.column_config.NumberColumn("Qtd", min_value=1, step=1, default=1),
                "Alt (cm)": st.column_config.NumberColumn("Alt (cm)", min_value=1.0, step=5.0, default=30.0),
                "Larg (cm)": st.column_config.NumberColumn("Larg (cm)", min_value=1.0, step=5.0, default=30.0),
                "Comp (cm)": st.column_config.NumberColumn("Comp (cm)", min_value=1.0, step=5.0, default=40.0),
                "Peso Total (kg)": st.column_config.NumberColumn("Peso Total (kg)", min_value=0.1, step=0.5, default=10.0, format="%.2f"),
            },
        )

        if edited_volumes is not None:
            # Save raw dict payload for history and API payloads (used in Retomar)
            st.session_state.volumes_data = edited_volumes.to_dict(orient="records")"""

if old_volume_block in content:
    content = content.replace(old_volume_block, new_volume_block)
else:
    print("WARNING: Could not find old volume block.")

# ==============================================================
# TICKET 1 FIX - RESUME QUOTE: we must reset volumes_df explicitly
# ==============================================================
old_resume_block = """                st.session_state.tipo_frete_input = item.get("tipo_frete", "FOB (Pago pelo Destinatário)")
                st.session_state.valor_nf_input = float(item.get("valor_nf", 0.0))
                st.session_state.volumes_data = item.get("volumes_data", [])
                st.session_state.active_module = "Nova Cotação"
                break
        del st.session_state.acao_retomar
        st.rerun()"""

new_resume_block = """                st.session_state.tipo_frete_input = item.get("tipo_frete", "FOB (Pago pelo Destinatário)")
                st.session_state.valor_nf_input = float(item.get("valor_nf", 0.0))
                
                # Restaurar volumes com o state dataframe fixo
                restored_volumes = item.get("volumes_data", [])
                import pandas as pd
                st.session_state.volumes_df = pd.DataFrame(restored_volumes) if restored_volumes else pd.DataFrame(columns=["Qtd", "Alt (cm)", "Larg (cm)", "Comp (cm)", "Peso Total (kg)"])
                st.session_state.volumes_data = restored_volumes
                if "volumes_editor" in st.session_state:
                    del st.session_state["volumes_editor"]
                    
                st.session_state.active_module = "Nova Cotação"
                break
        del st.session_state.acao_retomar
        st.rerun()"""

if old_resume_block in content:
    content = content.replace(old_resume_block, new_resume_block)
else:
    print("WARNING: Could not find old resume block.")

with open("main.py", "w", encoding="utf-8") as f:
    f.write(content)
print("Done patching.")
