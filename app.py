import streamlit as st
import pandas as pd
from datetime import datetime
import os
import json
import re

EXCEL_FILE = "inventario_finanzas.xlsx"
CONFIG_FILE = "config.json"

st.set_page_config(page_title="Sistema Integral Bimoneda (Inventario, POS, Clientes, Créditos, Finanzas)", page_icon="📦", layout="wide")

def load_tasa():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("tasa_cambio", 36.50)
        except:
            return 36.50
    return 36.50

def save_tasa(tasa_val):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"tasa_cambio": tasa_val}, f, ensure_ascii=False, indent=4)

if "tasa_cambio" not in st.session_state:
    st.session_state["tasa_cambio"] = load_tasa()

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            pd.DataFrame(columns=["ID", "Código", "Nombre", "Descripción", "Proveedor", "Precio Costo ($)", "Precio Venta ($)", "Stock Actual"]).to_excel(writer, sheet_name="Inventario", index=False)
            pd.DataFrame(columns=["ID Factura", "Fecha", "Cliente", "Cédula/RIF", "Teléfono", "Dirección", "Detalle Productos", "Total ($)", "Total (Bs)", "Condición Pago", "Método de Pago", "Tasa Usada"]).to_excel(writer, sheet_name="Facturas", index=False)
            pd.DataFrame(columns=["ID Transacción", "Fecha", "Tipo", "Categoría", "Descripción", "Monto ($)", "Monto (Bs)", "Método de Pago", "Tasa Usada"]).to_excel(writer, sheet_name="Finanzas", index=False)
            pd.DataFrame(columns=["ID Crédito", "Cliente", "Cédula/RIF", "Fecha", "Productos y Cantidades", "Total Crédito ($)", "Estado", "Fecha de Pago"]).to_excel(writer, sheet_name="Créditos", index=False)
            pd.DataFrame(columns=["Cédula/RIF", "Nombre", "Teléfono", "Dirección", "Crédito Habilitado", "Fecha Registro", "Total Compras ($)"]).to_excel(writer, sheet_name="Clientes", index=False)

init_excel()

# --- BARRA LATERAL ---
st.sidebar.header("⚙️ Configuración de Tasa")
tasa_actual_guardada = load_tasa()

with st.sidebar.form("form_tasa"):
    nueva_tasa = st.number_input("Tasa de Cambio (Bs / $)", min_value=0.01, value=float(tasa_actual_guardada), step=0.10, format="%.2f")
    btn_guardar_tasa = st.form_submit_button("💾 Guardar Tasa")
    if btn_guardar_tasa:
        save_tasa(nueva_tasa)
        st.session_state["tasa_cambio"] = nueva_tasa
        st.success(f"¡Tasa guardada: 1 $ = {nueva_tasa:,.2f} Bs!")
        st.rerun()

tasa = st.session_state["tasa_cambio"]
st.sidebar.info(f"💡 Tasa Actual:\n1 $ = **{tasa:,.2f} Bs**")

st.title("📦 Sistema Integral de Inventario, POS, Créditos y Finanzas")
st.markdown("---")

tabs = st.tabs(["📊 Dashboard", "🛒 Punto de Venta & Facturas", "🤝 Créditos / Fiados", "💰 Finanzas & Caja", "📦 Inventario & Carga", "👤 Base de Clientes", "📜 Historial de Facturas"])

METODOS_PAGO = [
    "Punto de Venta (Bs)",
    "Pago Móvil (Bs)",
    "Transferencia (Bs)",
    "Efectivo (Bs)",
    "Efectivo ($)",
    "Transferencia Zelle ($)",
    "Otro"
]

CATEGORIAS_FINANZAS = [
    "Venta Contado",
    "Cobro / Abono Crédito",
    "Ingreso Caja Chica",
    "Inmobiliaria / Alquiler",
    "Servicios Públicos (Luz, Agua, Internet)",
    "Pago de Nómina / Sueldos",
    "Compra de Inventario / Proveedores",
    "Gastos Operativos / Mantenimiento",
    "Otro Ingreso",
    "Otro Egreso"
]

def load_data():
    inv = pd.read_excel(EXCEL_FILE, sheet_name="Inventario")
    
    try: facturas = pd.read_excel(EXCEL_FILE, sheet_name="Facturas")
    except: facturas = pd.DataFrame(columns=["ID Factura", "Fecha", "Cliente", "Cédula/RIF", "Teléfono", "Dirección", "Detalle Productos", "Total ($)", "Total (Bs)", "Condición Pago", "Método de Pago", "Tasa Usada"])
    
    try: finanzas = pd.read_excel(EXCEL_FILE, sheet_name="Finanzas")
    except: finanzas = pd.DataFrame(columns=["ID Transacción", "Fecha", "Tipo", "Categoría", "Descripción", "Monto ($)", "Monto (Bs)", "Método de Pago", "Tasa Usada"])
    
    try: creditos = pd.read_excel(EXCEL_FILE, sheet_name="Créditos")
    except: creditos = pd.DataFrame(columns=["ID Crédito", "Cliente", "Cédula/RIF", "Fecha", "Productos y Cantidades", "Total Crédito ($)", "Estado", "Fecha de Pago"])
    
    try: clientes = pd.read_excel(EXCEL_FILE, sheet_name="Clientes")
    except: clientes = pd.DataFrame(columns=["Cédula/RIF", "Nombre", "Teléfono", "Dirección", "Crédito Habilitado", "Fecha Registro", "Total Compras ($)"])

    # Limpieza estricta y formato de columnas de texto en Clientes
    for col in ["Cédula/RIF", "Nombre", "Teléfono", "Dirección", "Crédito Habilitado"]:
        if col not in clientes.columns: clientes[col] = "Sí" if col == "Crédito Habilitado" else ""
        clientes[col] = clientes[col].astype(str).str.replace(r'\.0$', '', regex=True).replace("nan", "").replace("None", "").str.strip()
    
    if "Total Compras ($)" not in clientes.columns:
        clientes["Total Compras ($)"] = 0.0
    else:
        clientes["Total Compras ($)"] = pd.to_numeric(clientes["Total Compras ($)"], errors='coerce').fillna(0.0)

    for col in ["Código", "Proveedor", "Descripción", "Nombre"]:
        if col not in inv.columns: inv[col] = ""
        if col == "Código":
            inv[col] = inv[col].astype(str).str.replace(r'\.0$', '', regex=True).replace("nan", "").replace("None", "").str.strip()
        else:
            inv[col] = inv[col].astype(str).replace("nan", "").replace("None", "").str.strip()
        
    if "Precio Costo ($)" not in inv.columns and "Precio Costo" in inv.columns:
        inv.rename(columns={"Precio Costo": "Precio Costo ($)", "Precio Venta": "Precio Venta ($)"}, inplace=True)
    
    if "Condición Pago" not in facturas.columns: facturas["Condición Pago"] = "Contado"
    
    if "Cédula/RIF" not in creditos.columns: creditos["Cédula/RIF"] = "V-GENERAL"
    if "Total Crédito ($)" not in creditos.columns:
        if "Total Crédito" in creditos.columns:
            creditos.rename(columns={"Total Crédito": "Total Crédito ($)"}, inplace=True)
        elif "Total ($)" in creditos.columns:
            creditos.rename(columns={"Total ($)": "Total Crédito ($)"}, inplace=True)
        elif "Monto ($)" in creditos.columns:
            creditos.rename(columns={"Monto ($)": "Total Crédito ($)"}, inplace=True)
        else:
            creditos["Total Crédito ($)"] = 0.0

    return inv, facturas, finanzas, creditos, clientes

def save_data(inv, facturas, finanzas, creditos, clientes):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        inv.to_excel(writer, sheet_name="Inventario", index=False)
        facturas.to_excel(writer, sheet_name="Facturas", index=False)
        finanzas.to_excel(writer, sheet_name="Finanzas", index=False)
        creditos.to_excel(writer, sheet_name="Créditos", index=False)
        clientes.to_excel(writer, sheet_name="Clientes", index=False)

inv, facturas, finanzas, creditos, clientes = load_data()

def clean_col(df, col_usd, col_old=None):
    if col_usd in df.columns: return pd.to_numeric(df[col_usd], errors='coerce').fillna(0)
    elif col_old and col_old in df.columns: return pd.to_numeric(df[col_old], errors='coerce').fillna(0)
    return pd.Series(0, index=df.index)

def html5_camera_scanner_component(key_prefix="cam"):
    return st.components.v1.html(
        f"""
        <div style="text-align: center; font-family: sans-serif; padding: 10px; border: 1px dashed #007bff; border-radius: 8px;">
            <p style="margin: 0 0 8px 0; font-weight: bold; color: #007bff;">📷 Escáner de Cámara Activo</p>
            <video id="webcam-preview" style="width: 100%; max-width: 320px; height: 180px; border-radius: 6px; background: #000;" autoplay playsinline></video>
            <p style="font-size: 12px; color: #666; margin-top: 5px;">Apunta la cámara al código de barras. Haz clic abajo tras enfocar:</p>
            <button onclick="captureCode()" style="background-color: #28a745; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; font-weight: bold;">📸 Capturar Código Escaneado</button>
        </div>
        <script>
            navigator.mediaDevices.getUserMedia({{ video: {{ facingMode: "environment" }} }})
            .then(function(stream) {{
                let video = document.getElementById('webcam-preview');
                video.srcObject = stream;
            }}).catch(function(err) {{ console.log("Cámara no disponible: " + err); }});
            
            function captureCode() {{
                let simulatedCode = prompt("Ingrese el código detectado por la cámara (o confirme):");
                if(simulatedCode) {{
                    window.parent.postMessage({{ type: 'streamlit:setComponentValue', value: simulatedCode }}, '*');
                }}
            }}
        </script>
        """, height=280
    )

def generar_html_factura(id_fac, fecha_act, condicion_pago, metodo_pago_factura, nombre_cli_val, ced_final, carrito_items, tot_usd, tot_bs, tasa_val):
    filas_h = "".join([f"<tr><td>{i['Nombre']}</td><td style='text-align:center;'>{i['Cantidad']}</td><td style='text-align:right;'>${i['Precio Unitario']:,.2f}</td><td style='text-align:right;'>${i['Subtotal']:,.2f}</td></tr>" for i in carrito_items])
    
    plantilla = """
    <div style="border: 2px solid #333; padding: 20px; border-radius: 8px; background: #fff; color: #000; font-family: sans-serif; max-width: 600px; margin: auto;">
        <h2 style="text-align:center; margin:0;">🧾 FACTURA DE VENTA #{id_fac:04d} ({condicion_pago_upper})</h2>
        <p style="text-align:center; margin:5px 0;"><strong>Fecha:</strong> {fecha_act} | <strong>Tasa:</strong> {tasa_val:,.2f} Bs/$</p>
        <hr>
        <p><strong>Cliente:</strong> {nombre_cli_val} | <strong>Cédula/RIF:</strong> {ced_final}</p>
        <p><strong>Condición:</strong> {condicion_pago} | <strong>Método de Pago:</strong> {metodo_pago_factura}</p>
        <table style="width:100%; border-collapse:collapse;" border="1">
            <tr style="background:#f2f2f2;"><th>Producto</th><th>Cant</th><th>P.Unit ($)</th><th>Subtotal ($)</th></tr>
            {filas_h}
        </table>
        <div style="text-align:right; margin-top:15px;">
            <p style="margin:2px 0;"><strong>TOTAL REF ($):</strong> ${tot_usd:,.2f}</p>
            <h3 style="margin:2px 0; color:#008000;">TOTAL BOLÍVARES: {tot_bs:,.2f} Bs</h3>
        </div>
        <div style="text-align:center; margin-top:20px;">
            <button onclick="window.print();" style="background:#007bff; color:#fff; border:none; padding:10px 20px; font-weight:bold; border-radius:5px; cursor:pointer;">🖨️ Imprimir Factura</button>
        </div>
    </div>
    """
    return plantilla.format(
        id_fac=id_fac,
        condicion_pago_upper=condicion_pago.upper(),
        fecha_act=fecha_act,
        tasa_val=tasa_val,
        nombre_cli_val=nombre_cli_val,
        ced_final=ced_final,
        condicion_pago=condicion_pago,
        metodo_pago_factura=metodo_pago_factura,
        filas_h=filas_h,
        tot_usd=tot_usd,
        tot_bs=tot_bs
    )

# --- TAB 1: DASHBOARD ---
with tabs[0]:
    st.subheader("📈 Resumen General del Negocio")
    costo_col = clean_col(inv, "Precio Costo ($)", "Precio Costo")
    stock_col = clean_col(inv, "Stock Actual")
    total_inventario_valor = (costo_col * stock_col).sum() if not inv.empty else 0
    
    tot_v_col = clean_col(facturas, "Total ($)")
    total_ventas = tot_v_col.sum() if not facturas.empty else 0
    
    monto_fin_col = clean_col(finanzas, "Monto ($)", "Monto")
    ingresos = finanzas[finanzas["Tipo"] == "Ingreso"][monto_fin_col.name].sum() if not finanzas.empty and "Tipo" in finanzas.columns else 0
    egresos = finanzas[finanzas["Tipo"] == "Egreso"][monto_fin_col.name].sum() if not finanzas.empty and "Tipo" in finanzas.columns else 0
    balance_neto = ingresos - egresos
    
    tot_cred_col = clean_col(creditos, "Total Crédito ($)", "Total Crédito")
    total_fiado_pendiente = creditos[creditos["Estado"] == "Pendiente"][tot_cred_col.name].sum() if not creditos.empty and "Estado" in creditos.columns else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Valor Inventario", f"${total_inventario_valor:,.2f}", delta=f"{total_inventario_valor * tasa:,.2f} Bs")
    c2.metric("Total Ingresos", f"${ingresos:,.2f}", delta=f"{ingresos * tasa:,.2f} Bs")
    c3.metric("Total Egresos (Gastos)", f"${egresos:,.2f}", delta=f"-{egresos * tasa:,.2f} Bs", delta_color="inverse")
    c4.metric("Balance Neto Caja", f"${balance_neto:,.2f}", delta=f"{balance_neto * tasa:,.2f} Bs")
    c5.metric("Por Cobrar (Fiados)", f"${total_fiado_pendiente:,.2f}", delta=f"{total_fiado_pendiente * tasa:,.2f} Bs")

# --- TAB 2: PUNTO DE VENTA (POS) ---
with tabs[1]:
    st.subheader("🛒 Punto de Venta y Emisión de Facturas (Contado / Crédito)")
    if inv.empty:
        st.warning("No hay productos en inventario.")
    else:
        if "carrito" not in st.session_state: st.session_state["carrito"] = []

        col_cli, col_prod = st.columns([1, 1])

        with col_cli:
            st.markdown("#### 👤 Datos del Cliente y Modalidad de Pago")
            
            cedula_input = st.text_input("Cédula / RIF:", placeholder="Ej: V-12345678", key="pos_cedula_v13").strip()

            cliente_encontrado = False
            nombre_cli_val = ""
            tlf_cli_val = ""
            dir_cli_val = ""
            credito_permitido = "Sí"

            if cedula_input and not clientes.empty:
                coinc = clientes[clientes["Cédula/RIF"].astype(str).str.upper() == cedula_input.upper()]
                if not coinc.empty:
                    c_row = coinc.iloc[0]
                    nombre_cli_val = str(c_row["Nombre"])
                    tlf_cli_val = str(c_row.get("Teléfono", ""))
                    dir_cli_val = str(c_row.get("Dirección", ""))
                    credito_permitido = str(c_row.get("Crédito Habilitado", "Sí"))
                    cliente_encontrado = True
                    st.success(f"🟢 Cliente Frecuente Encontrado: {nombre_cli_val}")

            if cedula_input and not cliente_encontrado:
                st.warning("⚠️ El cliente no está registrado. Puedes registrarlo aquí mismo para continuar con la venta:")
                with st.form("form_reg_rapido_v13"):
                    reg_nombre = st.text_input("Nombre y Apellido / Razón Social *")
                    reg_tlf = st.text_input("Teléfono (Opcional)")
                    reg_dir = st.text_input("Dirección (Opcional)")
                    reg_cred = st.selectbox("¿Habilitar Crédito?", ["Sí", "No"])
                    btn_registrar_rapido = st.form_submit_button("💾 Registrar e Habilitar Venta")
                    
                    if btn_registrar_rapido:
                        if reg_nombre.strip():
                            nuevo_cli = pd.DataFrame([{
                                "Cédula/RIF": cedula_input.upper(),
                                "Nombre": reg_nombre.strip(),
                                "Teléfono": reg_tlf.strip(),
                                "Dirección": reg_dir.strip(),
                                "Crédito Habilitado": reg_cred,
                                "Fecha Registro": datetime.now().strftime("%Y-%m-%d"),
                                "Total Compras ($)": 0.0
                            }])
                            clientes = pd.concat([clientes, nuevo_cli], ignore_index=True)
                            save_data(inv, facturas, finanzas, creditos, clientes)
                            st.success("¡Cliente registrado exitosamente! Ya puedes continuar.")
                            st.rerun()
                        else:
                            st.error("El nombre es obligatorio para el registro.")
            elif cliente_encontrado:
                st.text_input("Nombre y Apellido / Razón Social *", value=nombre_cli_val, disabled=True)
                st.text_input("Teléfono (Opcional)", value=tlf_cli_val, disabled=True)
                st.text_input("Dirección (Opcional)", value=dir_cli_val, disabled=True)
                if credito_permitido == "No":
                    st.error("⛔ Este cliente tiene el **Crédito Bloqueado / No Habilitado** en el sistema.")
            else:
                st.text_input("Nombre y Apellido / Razón Social *", value="", disabled=True, placeholder="Ingrese cédula arriba...")
                st.text_input("Teléfono (Opcional)", value="", disabled=True)
                st.text_input("Dirección (Opcional)", value="", disabled=True)
            
            st.markdown("---")
            col_m1, col_m2 = st.columns(2)
            condicion_pago = col_m1.radio("Condición de la Venta:", ["Contado", "Crédito (Fiado)"], horizontal=True)
            metodo_pago_factura = col_m2.selectbox("Método de Pago", METODOS_PAGO, key="pos_metodo_v13")

        with col_prod:
            st.markdown("#### 📷 Entrada y Validación de Productos")
            modo_escaneo = st.radio("Método para ingresar Código de Barras:", ["🔌 Lectora / Pistola Laser", "📱 Cámara Móvil / Dispositivo"], horizontal=True, key="mode_pos_v13")
            
            def procesar_codigo_barras():
                val_ingresado = str(st.session_state.get("input_codigo_laser_v13", "")).strip()
                if val_ingresado.endswith('.0'):
                    val_ingresado = val_ingresado[:-2]
                    
                if val_ingresado:
                    prod_enc = inv[(inv["Código"].astype(str).str.lower() == val_ingresado.lower()) | (inv["Nombre"].astype(str).str.lower().str.contains(val_ingresado.lower()))]
                    if not prod_enc.empty:
                        p_row = prod_enc.iloc[0]
                        p_name = p_row["Nombre"]
                        p_price = float(p_row["Precio Venta ($)"])
                        p_stk = int(p_row["Stock Actual"])
                        
                        if p_stk > 0:
                            in_c = False
                            for item in st.session_state["carrito"]:
                                if item["Nombre"] == p_name:
                                    if item["Cantidad"] < p_stk:
                                        item["Cantidad"] += 1
                                        item["Subtotal"] = item["Cantidad"] * item["Precio Unitario"]
                                        st.session_state["ultimo_producto_agregado"] = f"✅ Sumado 1x **{p_name}** al carrito."
                                    else:
                                        st.session_state["ultimo_producto_agregado"] = f"⚠️ Stock máximo alcanzado para **{p_name}**."
                                    in_c = True
                                    break
                            if not in_c:
                                st.session_state["carrito"].append({"Nombre": p_name, "Precio Unitario": p_price, "Cantidad": 1, "Subtotal": p_price})
                                st.session_state["ultimo_producto_agregado"] = f"✅ Registrado y agregado: **{p_name}** (Precio: ${p_price:,.2f})"
                        else:
                            st.session_state["ultimo_producto_agregado"] = f"❌ El producto **{p_name}** está agotado."
                    else:
                        st.session_state["ultimo_producto_agregado"] = f"❌ No se encontró ningún producto con el código o nombre: '{val_ingresado}'"
                st.session_state["input_codigo_laser_v13"] = ""

            codigo_detectado = ""
            if modo_escaneo == "🔌 Lectora / Pistola Laser":
                st.text_input("Pase el láser por el código de barras o escriba y presione Enter:", placeholder="Escanear aquí y presionar Enter...", key="input_codigo_laser_v13", on_change=procesar_codigo_barras)
            else:
                st.caption("Permite acceder a la cámara del dispositivo móvil para leer el código:")
                html5_camera_scanner_component(key_prefix="pos_cam_v13")
                codigo_detectado = st.text_input("Código capturado con cámara:", placeholder="Código leído...", key="cam_code_pos_v13").strip()
                if codigo_detectado:
                    st.session_state["input_codigo_laser_v13"] = codigo_detectado
                    procesar_codigo_barras()

            if "ultimo_producto_agregado" in st.session_state and st.session_state["ultimo_producto_agregado"]:
                st.info(st.session_state["ultimo_producto_agregado"])

            st.markdown("---")
            prod_manual = st.selectbox("O seleccionar producto manualmente:", inv["Nombre"].tolist(), key="sel_manual_pos_v13")
            cant_manual = st.number_input("Cantidad", min_value=1, value=1, key="cant_manual_pos_v13")
            if st.button("➕ Agregar al Carrito Manualmente", key="btn_add_cart_v13"):
                p_row = inv[inv["Nombre"] == prod_manual].iloc[0]
                p_name = p_row["Nombre"]
                p_price = float(p_row["Precio Venta ($)"])
                p_stk = int(p_row["Stock Actual"])
                
                if p_stk >= cant_manual:
                    in_c = False
                    for item in st.session_state["carrito"]:
                        if item["Nombre"] == p_name:
                            if (item["Cantidad"] + cant_manual) <= p_stk:
                                item["Cantidad"] += cant_manual
                                item["Subtotal"] = item["Cantidad"] * item["Precio Unitario"]
                                st.success(f"Cantidad actualizada para {p_name}")
                            in_c = True
                            break
                    if not in_c:
                        st.session_state["carrito"].append({"Nombre": p_name, "Precio Unitario": p_price, "Cantidad": cant_manual, "Subtotal": p_price * cant_manual})
                        st.success(f"Agregado {cant_manual}x {p_name}")
                    st.rerun()

        st.markdown("---")
        st.markdown("#### 🛒 Detalle de Factura Actual")
        if len(st.session_state["carrito"]) > 0:
            df_c = pd.DataFrame(st.session_state["carrito"])
            st.dataframe(df_c, use_container_width=True)
            
            tot_usd = df_c["Subtotal"].sum()
            tot_bs = tot_usd * tasa
            
            c_tot1, c_tot2 = st.columns(2)
            c_tot1.metric("Total USD", f"${tot_usd:,.2f}")
            c_tot2.metric("Total Bolívares (Bs)", f"{tot_bs:,.2f} Bs")
            
            b1, b2 = st.columns(2)
            if b1.button("🗑️ Vaciar Carrito"):
                st.session_state["carrito"] = []
                st.session_state["ultimo_producto_agregado"] = ""
                st.rerun()
                
            if b2.button("🖨️ Procesar Venta y Emitir Factura"):
                if not cliente_encontrado:
                    st.error("❌ El cliente debe estar registrado para poder procesar la venta.")
                elif condicion_pago == "Crédito (Fiado)" and credito_permitido == "No":
                    st.error("⛔ Operación denegada: Este cliente no tiene autorizado el crédito en el sistema.")
                else:
                    ced_final = cedula_input.upper()
                    
                    exist_cli = clientes[clientes["Cédula/RIF"].astype(str).str.upper() == ced_final]
                    if not exist_cli.empty:
                        idx_c = exist_cli.index[0]
                        curr_tot = float(clientes.loc[idx_c, "Total Compras ($)"]) if not pd.isna(clientes.loc[idx_c, "Total Compras ($)"]) else 0.0
                        clientes.loc[idx_c, "Total Compras ($)"] = curr_tot + tot_usd

                    detalle_str = " | ".join([f"{i['Cantidad']}x {i['Nombre']} (${i['Precio Unitario']} c/u)" for i in st.session_state["carrito"]])
                    for item in st.session_state["carrito"]:
                        inv.loc[inv["Nombre"] == item["Nombre"], "Stock Actual"] -= item["Cantidad"]

                    id_fac = int(facturas["ID Factura"].max() + 1) if not facturas.empty and not pd.isna(facturas["ID Factura"].max()) else 1
                    fecha_act = datetime.now().strftime("%Y-%m-%d %H:%M")
                    
                    nueva_fac = pd.DataFrame([{"ID Factura": id_fac, "Fecha": fecha_act, "Cliente": nombre_cli_val, "Cédula/RIF": ced_final, "Teléfono": tlf_cli_val, "Dirección": dir_cli_val, "Detalle Productos": detalle_str, "Total ($)": tot_usd, "Total (Bs)": tot_bs, "Condición Pago": condicion_pago, "Método de Pago": metodo_pago_factura, "Tasa Usada": tasa}])
                    facturas = pd.concat([facturas, nueva_fac], ignore_index=True)

                    if condicion_pago == "Contado":
                        id_fin = int(finanzas["ID Transacción"].max() + 1) if not finanzas.empty and not pd.isna(finanzas["ID Transacción"].max()) else 1
                        nueva_fin = pd.DataFrame([{"ID Transacción": id_fin, "Fecha": fecha_act, "Tipo": "Ingreso", "Categoría": "Venta Contado", "Descripción": f"Factura #{id_fac:04d} - {nombre_cli_val}", "Monto ($)": tot_usd, "Monto (Bs)": tot_bs, "Método de Pago": metodo_pago_factura, "Tasa Usada": tasa}])
                        finanzas = pd.concat([finanzas, nueva_fin], ignore_index=True)
                    else:
                        exist_cred = creditos[(creditos["Cliente"].str.lower() == nombre_cli_val.lower()) & (creditos["Estado"] == "Pendiente")]
                        col_tot_name = "Total Crédito ($)" if "Total Crédito ($)" in creditos.columns else "Total Crédito"
                        if not exist_cred.empty:
                            idx_cr = exist_cred.index[0]
                            prod_prev = str(creditos.loc[idx_cr, "Productos y Cantidades"]) + " | " + detalle_str
                            tot_prev = float(creditos.loc[idx_cr, col_tot_name]) + tot_usd
                            creditos.loc[idx_cr, col_tot_name] = tot_prev
                            creditos.loc[idx_cr, "Productos y Cantidades"] = prod_prev
                        else:
                            id_cr = int(creditos["ID Crédito"].max() + 1) if not creditos.empty and not pd.isna(creditos["ID Crédito"].max()) else 1
                            nuevo_cr = pd.DataFrame([{
                                "ID Crédito": id_cr,
                                "Cliente": nombre_cli_val,
                                "Cédula/RIF": ced_final,
                                "Fecha": fecha_act,
                                "Productos y Cantidades": detalle_str,
                                "Total Crédito ($)": tot_usd,
                                "Estado": "Pendiente",
                                "Fecha de Pago": "Pendiente"
                            }])
                            creditos = pd.concat([creditos, nuevo_cr], ignore_index=True)

                    save_data(inv, facturas, finanzas, creditos, clientes)

                    html_fac = generar_html_factura(id_fac, fecha_act, condicion_pago, metodo_pago_factura, nombre_cli_val, ced_final, st.session_state["carrito"], tot_usd, tot_bs, tasa)
                    st.success(f"¡Venta a **{condicion_pago}** procesada con éxito! Factura #{id_fac:04d} generada.")
                    
                    st.download_button(
                        label="📥 Descargar Factura (HTML para enviar/imprimir)",
                        data=html_fac,
                        file_name=f"Factura_{id_fac:04d}_{nombre_cli_val.replace(' ', '_')}.html",
                        mime="text/html"
                    )

                    st.components.v1.html(html_fac, height=450, scrolling=True)
                    st.session_state["carrito"] = []
                    st.session_state["ultimo_producto_agregado"] = ""

# --- TAB 3: CRÉDITOS / FIADOS ---
with tabs[2]:
    st.subheader("🤝 Módulo de Gestión de Créditos, Abonos y Cuentas por Cobrar")
    
    sub_c1, sub_c2, sub_c3, sub_c4 = st.tabs(["📋 Administrar Cuentas y Abonos", "➕ Otorgar / Agregar Crédito", "✏️ Modificar Crédito", "🗑️ Eliminar Registro"])
    
    col_tot_name = "Total Crédito ($)" if "Total Crédito ($)" in creditos.columns else ("Total Crédito" if "Total Crédito" in creditos.columns else "Total ($)")

    with sub_c1:
        st.markdown("#### 💳 Cuentas Activas y Registro de Pagos")
        if not creditos.empty:
            pendientes = creditos[creditos["Estado"] == "Pendiente"]
            if pendientes.empty:
                st.success("🎉 ¡Excelente! No hay cuentas por cobrar pendientes en este momento.")
            else:
                for idx, row in pendientes.iterrows():
                    deuda_usd = float(row.get(col_tot_name, 0.0))
                    deuda_bs = deuda_usd * tasa
                    ced_val = row.get("Cédula/RIF", "V-GENERAL")
                    
                    with st.expander(f"👤 {row['Cliente']} ({ced_val}) | 💳 Deuda: ${deuda_usd:,.2f} (~ {deuda_bs:,.2f} Bs)"):
                        st.write("**Detalle de Compras y Abonos:**")
                        for item in str(row['Productos y Cantidades']).split(" | "):
                            st.write(f"• {item}")
                            
                        st.markdown("---")
                        col_ab, col_co = st.columns(2)
                        
                        with col_ab:
                            st.markdown("##### 💵 Registrar Abono")
                            moneda_abono = st.radio(f"Moneda:", ["Bolívares (Bs)", "Dólares ($)"], horizontal=True, key=f"m_ab_v13_{row['ID Crédito']}")
                            metodo_abono = st.selectbox("Método de Pago", METODOS_PAGO, key=f"met_ab_v13_{row['ID Crédito']}")
                            
                            if moneda_abono == "Bolívares (Bs)":
                                monto_ab_bs = st.number_input("Monto en Bs", min_value=0.01, value=float(min(deuda_bs, 100.0)), key=f"v_ab_bs_v13_{row['ID Crédito']}")
                                monto_ab_usd = monto_ab_bs / tasa if tasa > 0 else 0
                                st.caption(f"Descuenta de deuda: **${monto_ab_usd:,.2f}**")
                            else:
                                monto_ab_usd = st.number_input("Monto en $", min_value=0.01, max_value=max(deuda_usd, 0.01), value=float(min(deuda_usd, 5.0)), key=f"v_ab_usd_v13_{row['ID Crédito']}")
                                monto_ab_bs = monto_ab_usd * tasa
                                
                            if st.button("🔴 Registrar Abono", key=f"btn_ab_v13_{row['ID Crédito']}"):
                                nuevo_saldo = deuda_usd - monto_ab_usd
                                str_ab = f"💵 ABONO: -${monto_ab_usd:,.2f} (-{monto_ab_bs:,.2f} Bs) [{datetime.now().strftime('%d/%m %H:%M')}] ({metodo_abono})"
                                creditos.loc[idx, "Productos y Cantidades"] = str(row['Productos y Cantidades']) + " | " + str_ab
                                creditos.loc[idx, col_tot_name] = max(0.0, nuevo_saldo)
                                
                                if nuevo_saldo <= 0.001:
                                    creditos.loc[idx, "Estado"] = "Cancelado"
                                    creditos.loc[idx, "Fecha de Pago"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                                    st.balloons()
                                
                                id_f = int(finanzas["ID Transacción"].max() + 1) if not finanzas.empty and not pd.isna(finanzas["ID Transacción"].max()) else 1
                                new_ing = pd.DataFrame([{
                                    "ID Transacción": id_f,
                                    "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                    "Tipo": "Ingreso",
                                    "Categoría": "Cobro / Abono Crédito",
                                    "Descripción": f"Abono de {row['Cliente']}",
                                    "Monto ($)": monto_ab_usd,
                                    "Monto (Bs)": monto_ab_bs,
                                    "Método de Pago": metodo_abono,
                                    "Tasa Usada": tasa
                                }])
                                finanzas = pd.concat([finanzas, new_ing], ignore_index=True)
                                save_data(inv, facturas, finanzas, creditos, clientes)
                                st.success("¡Abono registrado con éxito!")
                                st.rerun()

                        with col_co:
                            st.markdown("##### 🏁 Liquidar Deuda Total")
                            st.write(f"Saldar total: **${deuda_usd:,.2f}** (**{deuda_bs:,.2f} Bs**)")
                            metodo_tot = st.selectbox("Método Liquidación", METODOS_PAGO, key=f"met_tot_v13_{row['ID Crédito']}")
                            
                            if st.button("✅ Saldar Deuda Completa", key=f"btn_saldar_v13_{row['ID Crédito']}"):
                                creditos.loc[idx, "Estado"] = "Cancelado"
                                creditos.loc[idx, "Fecha de Pago"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                                creditos.loc[idx, col_tot_name] = 0.0
                                
                                id_f = int(finanzas["ID Transacción"].max() + 1) if not finanzas.empty and not pd.isna(finanzas["ID Transacción"].max()) else 1
                                new_ing = pd.DataFrame([{
                                    "ID Transacción": id_f,
                                    "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                    "Tipo": "Ingreso",
                                    "Categoría": "Cobro / Abono Crédito",
                                    "Descripción": f"Cobro total a {row['Cliente']}",
                                    "Monto ($)": deuda_usd,
                                    "Monto (Bs)": deuda_bs,
                                    "Método de Pago": metodo_tot,
                                    "Tasa Usada": tasa
                                }])
                                finanzas = pd.concat([finanzas, new_ing], ignore_index=True)
                                save_data(inv, facturas, finanzas, creditos, clientes)
                                st.success("Cuenta saldada por completo.")
                                st.rerun()

            st.markdown("---")
            st.markdown("#### Historial General de Créditos")
            st.dataframe(creditos, use_container_width=True)
        else:
            st.info("No hay registros en la base de datos de créditos.")

    with sub_c2:
        st.markdown("#### ➕ Otorgar Crédito Manual a Cliente")
        with st.form("form_add_cred_v13"):
            c_cli1, c_cli2 = st.columns(2)
            c_nombre = c_cli1.text_input("Nombre del Cliente *")
            c_cedula = c_cli2.text_input("Cédula / RIF (Opcional)", value="V-GENERAL")
            
            c_prod1, c_prod2 = st.columns(2)
            c_prod_sel = c_prod1.selectbox("Seleccionar Producto de Inventario", inv["Nombre"].tolist() if not inv.empty else ["N/A"])
            c_cant = c_prod2.number_input("Cantidad", min_value=1, value=1)
            
            c_prec = st.number_input("Precio Unitario ($)", min_value=0.0, value=float(inv[inv["Nombre"]==c_prod_sel].iloc[0]["Precio Venta ($)"]) if not inv.empty else 0.0)
            c_obs = st.text_input("Observaciones / Notas")
            
            if st.form_submit_button("Guardar Registro de Crédito"):
                if c_nombre.strip():
                    monto_tot_cred = c_cant * c_prec
                    if not inv.empty and c_prod_sel in inv["Nombre"].values:
                        inv.loc[inv["Nombre"] == c_prod_sel, "Stock Actual"] -= c_cant
                        
                    det_str = f"• {c_cant}x {c_prod_sel} (${c_prec:,.2f} c/u) [{datetime.now().strftime('%d/%m %H:%M')}]"
                    if c_obs: det_str += f" ({c_obs})"
                    
                    exist_cr = creditos[(creditos["Cliente"].str.lower() == c_nombre.strip().lower()) & (creditos["Estado"] == "Pendiente")]
                    if not exist_cr.empty:
                        idx_cr = exist_cr.index[0]
                        prod_act = str(creditos.loc[idx_cr, "Productos y Cantidades"]) + " | " + det_str
                        tot_act = float(creditos.loc[idx_cr, col_tot_name]) + monto_tot_cred
                        creditos.loc[idx_cr, "Productos y Cantidades"] = prod_act
                        creditos.loc[idx_cr, col_tot_name] = tot_act
                    else:
                        id_cr = int(creditos["ID Crédito"].max() + 1) if not creditos.empty and not pd.isna(creditos["ID Crédito"].max()) else 1
                        nuevo_cr = pd.DataFrame([{
                            "ID Crédito": id_cr,
                            "Cliente": c_nombre.strip(),
                            "Cédula/RIF": c_cedula.strip().upper(),
                            "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                            "Productos y Cantidades": det_str,
                            "Total Crédito ($)": monto_tot_cred,
                            "Estado": "Pendiente",
                            "Fecha de Pago": "Pendiente"
                        }])
                        creditos = pd.concat([creditos, nuevo_cr], ignore_index=True)
                        
                    save_data(inv, facturas, finanzas, creditos, clientes)
                    st.success(f"¡Crédito otorgado con éxito a **{c_nombre}**!")
                    st.rerun()
                else:
                    st.error("El nombre del cliente es obligatorio.")

    with sub_c3:
        st.markdown("#### ✏️ Modificar Registro de Crédito")
        if not creditos.empty:
            sel_cred_edit = st.selectbox("Seleccione el Registro de Crédito a Modificar:", creditos.apply(lambda r: f"ID #{r['ID Crédito']} - {r['Cliente']} (${r.get(col_tot_name, 0.0):,.2f}) [{r['Estado']}]", axis=1).tolist(), key="sel_cred_edit_v13")
            
            id_selected = int(sel_cred_edit.split("#")[1].split(" ")[0])
            idx_cred = creditos[creditos["ID Crédito"] == id_selected].index[0]
            row_cr = creditos.loc[idx_cred]
            
            with st.form("form_edit_cred_v13"):
                ce_col1, ce_col2 = st.columns(2)
                e_cliente = ce_col1.text_input("Nombre del Cliente", value=str(row_cr["Cliente"]))
                e_cedula = ce_col2.text_input("Cédula / RIF", value=str(row_cr.get("Cédula/RIF", "V-GENERAL")))
                
                e_detalle = st.text_area("Detalle de Productos y Abonos", value=str(row_cr["Productos y Cantidades"]))
                
                ce_col3, ce_col4 = st.columns(2)
                e_monto = ce_col3.number_input("Monto Total de Deuda Pendiente ($)", min_value=0.0, value=float(row_cr.get(col_tot_name, 0.0)))
                e_estado = ce_col4.selectbox("Estado de la Cuenta", ["Pendiente", "Cancelado"], index=0 if row_cr["Estado"]=="Pendiente" else 1)
                
                if st.form_submit_button("💾 Guardar Cambios en Crédito"):
                    creditos.loc[idx_cred, "Cliente"] = e_cliente
                    creditos.loc[idx_cred, "Cédula/RIF"] = e_cedula
                    creditos.loc[idx_cred, "Productos y Cantidades"] = e_detalle
                    creditos.loc[idx_cred, col_tot_name] = e_monto
                    creditos.loc[idx_cred, "Estado"] = e_estado
                    if e_estado == "Cancelado" and row_cr["Estado"] != "Cancelado":
                        creditos.loc[idx_cred, "Fecha de Pago"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                    
                    save_data(inv, facturas, finanzas, creditos, clientes)
                    st.success("¡Registro de crédito modificado exitosamente!")
                    st.rerun()
        else:
            st.info("No hay créditos para editar.")

    with sub_c4:
        st.markdown("#### 🗑️ Eliminar Registro de Crédito")
        if not creditos.empty:
            sel_cred_del = st.selectbox("Seleccione el Registro de Crédito que desea Eliminar:", creditos.apply(lambda r: f"ID #{r['ID Crédito']} - {r['Cliente']} (${r.get(col_tot_name, 0.0):,.2f}) [{r['Estado']}]", axis=1).tolist(), key="sel_cred_del_v13")
            
            id_del_selected = int(sel_cred_del.split("#")[1].split(" ")[0])
            st.warning(f"⚠️ ¿Desea borrar definitivamente la cuenta ID #{id_del_selected}? Se eliminará del registro general.")
            
            if st.button("🗑️ Confirmar Eliminar Crédito"):
                creditos = creditos[creditos["ID Crédito"] != id_del_selected]
                save_data(inv, facturas, finanzas, creditos, clientes)
                st.success(f"Registro de crédito ID #{id_del_selected} eliminado correctamente.")
                st.rerun()
        else:
            st.info("No hay créditos para eliminar.")

# --- TAB 4: FINANZAS & CAJA ---
with tabs[3]:
    st.subheader("💰 Control Financiero, Gastos e Ingresos Manuales")
    
    sub_f1, sub_f2, sub_f3 = st.tabs(["📊 Balance y Movimientos", "➕ Registrar Ingreso / Egreso Manual", "🗑️ Eliminar Movimiento"])
    
    with sub_f1:
        st.markdown("#### 📋 Historial Completo de Caja y Finanzas")
        if not finanzas.empty:
            m_col = clean_col(finanzas, "Monto ($)", "Monto")
            tot_ing = finanzas[finanzas["Tipo"] == "Ingreso"][m_col.name].sum() if "Tipo" in finanzas.columns else 0
            tot_egr = finanzas[finanzas["Tipo"] == "Egreso"][m_col.name].sum() if "Tipo" in finanzas.columns else 0
            bal = tot_ing - tot_egr
            
            f_col1, f_col2, f_col3 = st.columns(3)
            f_col1.metric("Ingresos Totales", f"${tot_ing:,.2f}", delta=f"{tot_ing * tasa:,.2f} Bs")
            f_col2.metric("Egresos Totales (Gastos)", f"${tot_egr:,.2f}", delta=f"-{tot_egr * tasa:,.2f} Bs", delta_color="inverse")
            f_col3.metric("Balance Neto", f"${bal:,.2f}", delta=f"{bal * tasa:,.2f} Bs")
            
            st.markdown("---")
            st.dataframe(finanzas, use_container_width=True)
        else:
            st.info("No hay movimientos registrados en finanzas.")
            
    with sub_f2:
        st.markdown("#### ➕ Registrar Movimiento Manual (Inmobiliaria, Caja Chica, Gastos, etc.)")
        with st.form("form_add_fin_v13"):
            col_t1, col_t2 = st.columns(2)
            tipo_mov = col_t1.radio("Tipo de Movimiento *", ["Ingreso", "Egreso"], horizontal=True)
            categoria_mov = col_t2.selectbox("Categoría / Motivo *", CATEGORIAS_FINANZAS)
            
            col_d1, col_d2 = st.columns(2)
            desc_mov = col_d1.text_input("Descripción / Observación *", placeholder="Ej: Pago de alquiler de local, Reposición de caja chica...")
            metodo_mov = col_d2.selectbox("Método de Pago Utilizado *", METODOS_PAGO, key="fin_metodo_manual_v13")
            
            col_m1, col_m2 = st.columns(2)
            moneda_mov = col_m1.radio("Moneda del Monto *", ["Dólares ($)", "Bolívares (Bs)"], horizontal=True)
            
            if moneda_mov == "Dólares ($)":
                monto_usd_in = col_m2.number_input("Monto en USD ($)", min_value=0.01, value=10.0, step=1.0)
                monto_bs_in = monto_usd_in * tasa
                st.caption(f"Equivalente a tasa actual: **{monto_bs_in:,.2f} Bs**")
            else:
                monto_bs_in = col_m2.number_input("Monto en Bolívares (Bs)", min_value=0.01, value=100.0, step=10.0)
                monto_usd_in = monto_bs_in / tasa if tasa > 0 else 0
                st.caption(f"Equivalente a tasa actual: **${monto_usd_in:,.2f} USD**")

            if st.form_submit_button("💾 Guardar Movimiento Financiero"):
                if desc_mov.strip():
                    id_f = int(finanzas["ID Transacción"].max() + 1) if not finanzas.empty and not pd.isna(finanzas["ID Transacción"].max()) else 1
                    fecha_act = datetime.now().strftime("%Y-%m-%d %H:%M")
                    
                    nuevo_mov = pd.DataFrame([{
                        "ID Transacción": id_f,
                        "Fecha": fecha_act,
                        "Tipo": tipo_mov,
                        "Categoría": categoria_mov,
                        "Descripción": desc_mov.strip(),
                        "Monto ($)": monto_usd_in,
                        "Monto (Bs)": monto_bs_in,
                        "Método de Pago": metodo_mov,
                        "Tasa Usada": tasa
                    }])
                    finanzas = pd.concat([finanzas, nuevo_mov], ignore_index=True)
                    save_data(inv, facturas, finanzas, creditos, clientes)
                    st.success(f"¡{tipo_mov} registrado exitosamente por **${monto_usd_in:,.2f}** ({categoria_mov})!")
                    st.rerun()
                else:
                    st.error("Por favor ingrese una descripción para el movimiento.")

    with sub_f3:
        st.markdown("#### 🗑️ Eliminar Movimiento Financiero")
        if not finanzas.empty:
            sel_fin_del = st.selectbox("Seleccione la Transacción a Eliminar:", finanzas.apply(lambda r: f"ID #{r['ID Transacción']} - [{r['Tipo']}] {r['Categoría']} - ${r['Monto ($)']:,.2f} ({r['Fecha']})", axis=1).tolist(), key="sel_fin_del_v13")
            
            id_fin_selected = int(sel_fin_del.split("#")[1].split(" ")[0])
            st.warning(f"⚠️ ¿Desea borrar la transacción ID #{id_fin_selected}? Esto actualizará los balances del negocio.")
            
            if st.button("🗑️ Confirmar Eliminar Transacción"):
                finanzas = finanzas[finanzas["ID Transacción"] != id_fin_selected]
                save_data(inv, facturas, finanzas, creditos, clientes)
                st.success(f"Transacción ID #{id_fin_selected} eliminada correctamente.")
                st.rerun()
        else:
            st.info("No hay transacciones para eliminar.")

# --- TAB 5: INVENTARIO & CARGA ---
with tabs[4]:
    st.subheader("📦 Control de Inventario y Gestión de Productos")
    sub_inv1, sub_inv2, sub_inv3, sub_inv4 = st.tabs(["➕ Agregar Manual", "✏️ Modificar / Editar", "🗑️ Eliminar Producto", "📥 Importar Masivo desde Excel"])
    
    with sub_inv1:
        st.markdown("#### ➕ Agregar Nuevo Producto")
        modo_bar_inv = st.radio("¿Cómo deseas ingresar el código de barras?", ["🔌 Lectora / Pistola Laser", "📱 Cámara Móvil / Webcam"], horizontal=True, key="mode_add_inv_v13")
        cod_autodetect = ""
        if modo_bar_inv == "📱 Cámara Móvil / Webcam":
            html5_camera_scanner_component(key_prefix="inv_add_cam_v13")
            cod_autodetect = st.text_input("Código capturado por cámara:", key="cam_code_inv_add_v13").strip()

        with st.form("form_add_prod_v13"):
            c1, c2, c3 = st.columns(3)
            cod_in = c1.text_input("Código de Barras / SKU", value=cod_autodetect, placeholder="Escanee con pistola o escriba...")
            nom_in = c2.text_input("Nombre del Producto *")
            prov_in = c3.text_input("Proveedor")
            desc_in = st.text_input("Descripción / Categoría")
            
            c4, c5, c6 = st.columns(3)
            stk_in = c4.number_input("Stock Inicial", min_value=0, value=1)
            cost_in = c5.number_input("Precio Costo ($)", min_value=0.0, value=0.0)
            vent_in = c6.number_input("Precio Venta ($)", min_value=0.0, value=0.0)
            
            if st.form_submit_button("Guardar Nuevo Producto"):
                if nom_in.strip():
                    nid = int(inv["ID"].max() + 1) if not inv.empty and not pd.isna(inv["ID"].max()) else 1
                    cod_str = str(cod_in).strip()
                    if cod_str.endswith('.0'):
                        cod_str = cod_str[:-2]
                    cod_final = cod_str if cod_str else f"BAR-{nid:04d}"
                    nuevo_p = pd.DataFrame([{"ID": nid, "Código": cod_final, "Nombre": nom_in.strip(), "Descripción": desc_in, "Proveedor": prov_in, "Precio Costo ($)": cost_in, "Precio Venta ($)": vent_in, "Stock Actual": stk_in}])
                    inv = pd.concat([inv, nuevo_p], ignore_index=True)
                    save_data(inv, facturas, finanzas, creditos, clientes)
                    st.success(f"¡Producto **{nom_in}** guardado correctamente!")
                    st.rerun()
                else: st.error("El nombre del producto es obligatorio.")

    with sub_inv2:
        st.markdown("#### ✏️ Modificar o Editar Producto Existente")
        if not inv.empty:
            prod_edit_sel = st.selectbox("Seleccione el producto a editar:", inv["Nombre"].tolist(), key="sel_edit_inv_v13")
            row_e = inv[inv["Nombre"] == prod_edit_sel].iloc[0]
            with st.form("form_edit_prod_v13"):
                ec1, ec2, ec3 = st.columns(3)
                e_cod = ec1.text_input("Código de Barras", value=str(row_e.get("Código", "")))
                e_nom = ec2.text_input("Nombre del Producto", value=str(row_e["Nombre"]))
                e_prov = ec3.text_input("Proveedor", value=str(row_e.get("Proveedor", "")))
                e_desc = st.text_input("Descripción", value=str(row_e.get("Descripción", "")))
                
                ec4, ec5, ec6 = st.columns(3)
                e_stk = ec4.number_input("Stock Actual", min_value=0, value=int(row_e["Stock Actual"]))
                e_pc = ec5.number_input("Precio Costo ($)", min_value=0.0, value=float(row_e.get("Precio Costo ($)", 0)))
                e_pv = ec6.number_input("Precio Venta ($)", min_value=0.0, value=float(row_e.get("Precio Venta ($)", 0)))
                
                if st.form_submit_button("💾 Guardar Cambios"):
                    idx = inv[inv["Nombre"] == prod_edit_sel].index[0]
                    inv["Código"] = inv["Código"].astype(str)
                    clean_e_cod = str(e_cod).strip()
                    if clean_e_cod.endswith('.0'):
                        clean_e_cod = clean_e_cod[:-2]
                    inv.loc[idx, "Código"] = clean_e_cod
                    inv.loc[idx, "Nombre"] = str(e_nom)
                    inv.loc[idx, "Proveedor"] = str(e_prov)
                    inv.loc[idx, "Descripción"] = str(e_desc)
                    inv.loc[idx, "Stock Actual"] = e_stk
                    inv.loc[idx, "Precio Costo ($)"] = e_pc
                    inv.loc[idx, "Precio Venta ($)"] = e_pv
                    save_data(inv, facturas, finanzas, creditos, clientes)
                    st.success(f"¡Producto **{e_nom}** actualizado!")
                    st.rerun()
        else: st.info("No hay productos.")

    with sub_inv3:
        st.markdown("#### 🗑️ Eliminar Producto")
        if not inv.empty:
            prod_del_sel = st.selectbox("Seleccione producto a eliminar:", inv["Nombre"].tolist(), key="sel_del_inv_v13")
            if st.button("🗑️ Confirmar Eliminar"):
                inv = inv[inv["Nombre"] != prod_del_sel]
                save_data(inv, facturas, finanzas, creditos, clientes)
                st.success("Producto eliminado.")
                st.rerun()
        else: st.info("No hay productos.")

    with sub_inv4:
        st.markdown("#### 📥 Importar Masivo desde Excel")
        archivo_subido = st.file_uploader("Sube tu archivo Excel o CSV", type=["xlsx", "csv"])
        if archivo_subido is not None and st.button("🚀 Cargar Inventario"):
            df_imp = pd.read_csv(archivo_subido) if archivo_subido.name.endswith('.csv') else pd.read_excel(archivo_subido)
            for _, r in df_imp.iterrows():
                nom_p = str(r["Nombre"])
                if pd.isna(nom_p) or not nom_p.strip(): continue
                ex = inv[inv["Nombre"].str.lower() == nom_p.lower()]
                stk_val = int(r["Stock Actual"]) if "Stock Actual" in df_imp.columns and not pd.isna(r["Stock Actual"]) else 0
                pv_val = float(r["Precio Venta ($)"]) if "Precio Venta ($)" in df_imp.columns and not pd.isna(r["Precio Venta ($)"]) else 0.0
                
                cod_imp = str(r.get("Código", ""))
                if cod_imp.endswith('.0'):
                    cod_imp = cod_imp[:-2]
                
                if not ex.empty:
                    inv.loc[ex.index[0], "Stock Actual"] += stk_val
                else:
                    nid = int(inv["ID"].max() + 1) if not inv.empty and not pd.isna(inv["ID"].max()) else 1
                    final_code = cod_imp if cod_imp and cod_imp != "nan" else f"BAR-{nid:04d}"
                    inv = pd.concat([inv, pd.DataFrame([{"ID": nid, "Código": final_code, "Nombre": nom_p, "Descripción": str(r.get("Descripción", "")), "Proveedor": str(r.get("Proveedor", "")), "Precio Costo ($)": float(r.get("Precio Costo ($)", 0)), "Precio Venta ($)": pv_val, "Stock Actual": stk_val}])], ignore_index=True)
            save_data(inv, facturas, finanzas, creditos, clientes)
            st.success("¡Importación lista!")
            st.rerun()

    st.dataframe(inv, use_container_width=True)

# --- TAB 6: CLIENTES ---
with tabs[5]:
    st.subheader("👤 Base de Clientes Frecuentes y Gestión de Crédito")
    
    sub_cli1, sub_cli2 = st.tabs(["📋 Lista de Clientes", "✏️ Modificar Datos y Estatus de Crédito"])
    
    with sub_cli1:
        st.markdown("#### 📋 Clientes Registrados")
        if not clientes.empty:
            st.dataframe(clientes, use_container_width=True)
        else:
            st.info("No hay clientes registrados en el sistema.")

    with sub_cli2:
        st.markdown("#### ✏️ Modificar Información y Otorgar / Denegar Crédito")
        if not clientes.empty:
            sel_cliente_mod = st.selectbox(
                "Seleccione el cliente que desea modificar:",
                clientes.apply(lambda r: f"{r['Cédula/RIF']} - {r['Nombre']} (Crédito: {r.get('Crédito Habilitado', 'Sí')})", axis=1).tolist(),
                key="sel_mod_cliente_v13"
            )
            
            cedula_sel = sel_cliente_mod.split(" - ")[0]
            idx_cli = clientes[clientes["Cédula/RIF"].astype(str).str.upper() == cedula_sel.upper()].index[0]
            row_cli = clientes.loc[idx_cli]
            
            with st.form("form_modificar_cliente_v13"):
                mc1, mc2 = st.columns(2)
                m_nom = mc1.text_input("Nombre y Apellido / Razón Social", value=str(row_cli["Nombre"]))
                m_ced = mc2.text_input("Cédula / RIF", value=str(row_cli["Cédula/RIF"]))
                
                mc3, mc4 = st.columns(2)
                m_tlf = mc3.text_input("Teléfono", value=str(row_cli.get("Teléfono", "")))
                m_dir = mc4.text_input("Dirección", value=str(row_cli.get("Dirección", "")))
                
                estatus_actual = str(row_cli.get("Crédito Habilitado", "Sí"))
                idx_est = 0 if estatus_actual == "Sí" else 1
                m_credito_hab = st.selectbox("Estatus de Crédito (¿Permitir ventas a crédito / fiado?)", ["Sí", "No"], index=idx_est)
                
                if st.form_submit_button("💾 Guardar Cambios del Cliente"):
                    clientes.loc[idx_cli, "Nombre"] = m_nom.strip()
                    clientes.loc[idx_cli, "Cédula/RIF"] = m_ced.strip().upper()
                    clientes.loc[idx_cli, "Teléfono"] = m_tlf.strip()
                    clientes.loc[idx_cli, "Dirección"] = m_dir.strip()
                    clientes.loc[idx_cli, "Crédito Habilitado"] = m_credito_hab
                    
                    save_data(inv, facturas, finanzas, creditos, clientes)
                    st.success(f"¡Datos del cliente **{m_nom}** actualizados con éxito!")
                    st.rerun()
        else:
            st.info("No hay clientes para modificar.")

# --- TAB 7: HISTORIAL FACTURAS & REIMPRESIÓN ---
with tabs[6]:
    st.subheader("📜 Historial de Facturas y Reimpresión por Extravío")
    if not facturas.empty:
        facturas["Fecha_DT"] = pd.to_datetime(facturas["Fecha"], errors='coerce')
        hoy = datetime.now().date()
        filtro = st.radio("Periodo:", ["Diario (Hoy)", "Semanal (Últimos 7 días)", "Mensual (Mes Actual)", "Todos"], horizontal=True, key="filtro_hist_v13")
        df_h = facturas.copy()
        if filtro == "Diario (Hoy)": df_h = df_h[df_h["Fecha_DT"].dt.date == hoy]
        elif filtro == "Semanal (Últimos 7 días)": df_h = df_h[df_h["Fecha_DT"] >= (pd.Timestamp(hoy) - pd.Timedelta(days=7))]
        elif filtro == "Mensual (Mes Actual)": df_h = df_h[(df_h["Fecha_DT"].dt.year == hoy.year) & (df_h["Fecha_DT"].dt.month == hoy.month)]
        
        st.dataframe(df_h.drop(columns=["Fecha_DT"], errors='ignore'), use_container_width=True)
        
        st.markdown("---")
        st.markdown("#### 🖨️ Reimprimir Factura Seleccionada")
        
        lista_fac_opciones = list(df_h.apply(lambda r: f"Factura #{int(r['ID Factura']):04d} - {r['Cliente']} (${r['Total ($)']:,.2f}) [{r['Fecha']}]", axis=1))
        
        if lista_fac_opciones:
            fac_sel_reimp = st.selectbox("Seleccione la factura extraviada a reimprimir:", lista_fac_opciones, key="sel_reimpresion_fac_v13")
            
            match_id = re.search(r"#(\d+)", fac_sel_reimp)
            if match_id:
                id_f_reimp = int(match_id.group(1))
                row_fac = facturas[facturas["ID Factura"] == id_f_reimp].iloc[0]
                
                if st.button("🖨️ Generar Comprobante para Reimpresión"):
                    detalles_txt = str(row_fac["Detalle Productos"])
                    
                    plantilla_reimp = """
                    <div style="border: 2px solid #333; padding: 20px; border-radius: 8px; background: #fff; color: #000; font-family: sans-serif; max-width: 600px; margin: auto;">
                        <h2 style="text-align:center; margin:0;">🧾 REIMPRESIÓN FACTURA DE VENTA #{id_fac:04d}</h2>
                        <p style="text-align:center; margin:5px 0;"><strong>Fecha original:</strong> {fecha_orig} | <strong>Tasa:</strong> {tasa_val:,.2f} Bs/$</p>
                        <hr>
                        <p><strong>Cliente:</strong> {cliente} | <strong>Cédula/RIF:</strong> {cedula}</p>
                        <p><strong>Condición:</strong> {condicion} | <strong>Método de Pago:</strong> {metodo}</p>
                        <p><strong>Detalle de Productos:</strong></p>
                        <p style="background: #f9f9f9; padding: 10px; border-radius: 5px;">{detalles}</p>
                        <div style="text-align:right; margin-top:15px;">
                            <p style="margin:2px 0;"><strong>TOTAL REF ($):</strong> ${tot_usd:,.2f}</p>
                            <h3 style="margin:2px 0; color:#008000;">TOTAL BOLÍVARES: {tot_bs:,.2f} Bs</h3>
                        </div>
                        <div style="text-align:center; margin-top:20px;">
                            <button onclick="window.print();" style="background:#28a745; color:#fff; border:none; padding:10px 20px; font-weight:bold; border-radius:5px; cursor:pointer;">🖨️ Imprimir Reimpresión</button>
                        </div>
                    </div>
                    """
                    html_reimp = plantilla_reimp.format(
                        id_fac=int(row_fac['ID Factura']),
                        fecha_orig=row_fac['Fecha'],
                        tasa_val=row_fac.get('Tasa Usada', tasa),
                        cliente=row_fac['Cliente'],
                        cedula=row_fac.get('Cédula/RIF', 'N/A'),
                        condicion=row_fac.get('Condición Pago', 'Contado'),
                        metodo=row_fac.get('Método de Pago', 'N/A'),
                        detalles=detalles_txt,
                        tot_usd=row_fac['Total ($)'],
                        tot_bs=row_fac['Total (Bs)']
                    )
                    
                    st.success(f"Factura #{id_f_reimp:04d} lista para reimprimir:")
                    st.download_button(
                        label="📥 Descargar Factura Duplicada (HTML)",
                        data=html_reimp,
                        file_name=f"Reimpresion_Factura_{id_f_reimp:04d}.html",
                        mime="text/html"
                    )
                    st.components.v1.html(html_reimp, height=450, scrolling=True)
    else: 
        st.info("No hay facturas registradas todavía.")