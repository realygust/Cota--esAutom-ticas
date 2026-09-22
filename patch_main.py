import json
import re

with open("main.py", "r", encoding="utf-8") as f:
    content = f.read()

# Replace payload
old_payload = """            salvar_historico_item({
                "id": f"COT-{datetime.now().strftime('%Y%m%d-%H%M%S')}",
                "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "origem": f"{info_o.get('cidade','')}/{info_o.get('uf','')}".strip("/"),
                "destino": f"{cidade_dest}/{uf_dest}".strip("/"),
                "cep_origem": cep_origem,
                "cep_destino": cep_destino,
                "peso_kg": peso_real_total,
                "peso_tarifado_kg": peso_tarifado,
                "valor_nf": valor_nf,
                "melhor_transportadora": melhor_item["Transportadora"],
                "melhor_valor": melhor_item["Valor Frete (R$)"],
                "melhor_prazo": int(melhor_item.get("Prazo (Dias Úteis)") or 0),
                "melhor_tipo": melhor_item.get("Base Tarifária", "Tabela Contratual"),
            })"""

new_payload = """            salvar_historico_item({
                "id": f"COT-{datetime.now().strftime('%Y%m%d-%H%M%S')}",
                "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "status": "Pendente",
                "tipo_frete": st.session_state.get("tipo_frete_input", "FOB (Pago pelo Destinatário)"),
                "doc_destinatario": doc_destinatario,
                "volumes_data": st.session_state.get("volumes_data", []),
                "origem": f"{info_o.get('cidade','')}/{info_o.get('uf','')}".strip("/"),
                "destino": f"{cidade_dest}/{uf_dest}".strip("/"),
                "cep_origem": cep_origem,
                "cep_destino": cep_destino,
                "peso_kg": peso_real_total,
                "peso_tarifado_kg": peso_tarifado,
                "valor_nf": valor_nf,
                "melhor_transportadora": melhor_item["Transportadora"],
                "melhor_valor": melhor_item["Valor Frete (R$)"],
                "melhor_prazo": int(melhor_item.get("Prazo (Dias Úteis)") or 0),
                "melhor_tipo": melhor_item.get("Base Tarifária", "Tabela Contratual"),
            })"""
content = content.replace(old_payload, new_payload)

# Replace aba name
content = content.replace('"Histórico de Cotações (Em Breve)"', '"Histórico de Cotações"')

# Replace History section
old_history_start = """# MÓDULO 4: HISTÓRICO DE COTAÇÕES
# ==============================================================================
elif aba == "Histórico de Cotações":"""

new_history = """# MÓDULO 4: HISTÓRICO DE COTAÇÕES
# ==============================================================================
elif aba == "Histórico de Cotações":
    st.markdown("#### Histórico de Cotações")

    historico = carregar_historico()
    
    # Processar ações
    if "acao_aprovar" in st.session_state:
        cot_id = st.session_state.acao_aprovar
        for item in historico:
            if item.get("id") == cot_id:
                item["status"] = "Aprovada"
                break
        try:
            with open(HISTORICO_FILE, "w", encoding="utf-8") as f:
                json.dump(historico, f, ensure_ascii=False, indent=2)
        except Exception: pass
        del st.session_state.acao_aprovar
        st.rerun()

    if "acao_retomar" in st.session_state:
        cot_id = st.session_state.acao_retomar
        for item in historico:
            if item.get("id") == cot_id:
                st.session_state.doc_dest_input = item.get("doc_destinatario", "")
                st.session_state.cep_d_input = item.get("cep_destino", "")
                st.session_state.tipo_frete_input = item.get("tipo_frete", "FOB (Pago pelo Destinatário)")
                st.session_state.valor_nf_input = float(item.get("valor_nf", 0.0))
                st.session_state.volumes_data = item.get("volumes_data", [])
                st.session_state.aba_ativa = "Nova Cotação"
                break
        del st.session_state.acao_retomar
        st.rerun()

    # Filtrar TTL 48h
    historico_valido = []
    from datetime import datetime, timedelta
    agora = datetime.now()
    
    for item in historico:
        status = item.get("status", "Pendente")
        dh_str = item.get("data_hora", "")
        if status == "Pendente":
            try:
                dt_item = datetime.strptime(dh_str, "%d/%m/%Y %H:%M")
                if agora - dt_item > timedelta(hours=48):
                    continue
            except:
                pass
        historico_valido.append(item)

    if not historico_valido:
        st.info("Nenhuma cotação recente ou aprovada registrada. Faça uma cotação na aba 'Nova Cotação'.")
    else:
        for item in historico_valido:
            with st.container(border=True):
                st.markdown(f"**{item.get('data_hora')}** | **CNPJ Cliente:** {item.get('doc_destinatario', 'N/A')}")
                st.markdown(f"**Destino:** {item.get('destino')} | **Valor NF:** R$ {item.get('valor_nf', 0):,.2f}")
                st.markdown(f"**Melhor Oferta:** {item.get('melhor_transportadora')} - R$ {item.get('melhor_valor', 0):,.2f}")
                
                status = item.get('status', 'Pendente')
                cor_status = "green" if status == "Aprovada" else "orange"
                st.markdown(f"**Status:** :{cor_status}[{status}]")
                
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("Retomar Cotação", key=f"retomar_{item['id']}"):
                        st.session_state.acao_retomar = item["id"]
                        st.rerun()
                with c2:
                    if status == "Pendente":
                        if st.button("Aprovar", key=f"aprovar_{item['id']}"):
                            st.session_state.acao_aprovar = item["id"]
                            st.rerun()"""

parts = content.split('# MÓDULO 4: HISTÓRICO DE COTAÇÕES\n# ==============================================================================')
if len(parts) == 2:
    content = parts[0] + new_history
    with open("main.py", "w", encoding="utf-8") as f:
        f.write(content)
    print("Done")
else:
    print("Failed to find history section")
