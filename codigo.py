import streamlit as st
import pandas as pd
import numpy as np
import io

if 'df_original' not in st.session_state:
    st.session_state['df_original'] = None


# 1. Configuración de la página de Streamlit
st.set_page_config(
    page_title="Balance Financiero Proyectado",
    layout="centered"
)

# Renderizar logo y título
try:
    st.image("logo3.png", width="stretch")
except:
    pass

st.title("Balance Financiero Proyectado")

# 2. Función optimizada para cargar archivos pesados con caché
@st.cache_data
def cargar_datos(archivo_subido):
    return pd.read_excel(archivo_subido, sheet_name=0, engine='openpyxl')

def balance(df):     
    # Convertimos la columna a fechas reales
    df['FECHA'] = pd.to_datetime(df['FECHA_CONEX'], format='%d/%m/%Y', errors='coerce')
    df['ID'] = df['SUCURSAL_CVE'].astype(str) + df['UNIDAD_OP_CVE'].astype(str) + df['FOLIO_DEF'].astype(str)
    
    # Filtrados iniciales reduciendo el tamaño del dataframe lo antes posible
    df_limpio = df.dropna(subset=['FECHA'])
    df_limpio = df_limpio[df_limpio['ESTATUS'] == 'Conectado']
    df_limpio = df_limpio[df_limpio['TIPO_CEGAP'] != 'Cegap Nac. de Proveedores Descontados']

    # Identificar la máscara de duplicados recurrentes para excluir
    mask_duplicado = (
        df_limpio['FOLIO_DEF'].between(49999, 79999) &
        df_limpio['TIPO_CEGAP'].isin(['Cegap de Erogacion', 'Cegap de Erogacion de Proveedores (Mercancias)']) &
        (df_limpio['TIPO_PAGO'] == 'Cegap de registro')
    )
    
    base = df_limpio[~mask_duplicado].copy()

    # --- SECCIÓN GASTOS ---
    gasto = base[
        (base['CAPITULO'] != 7000000) &
        (~base['PARTIDA'].isin([43101001, 43701001, 39908031, 39908046])) &
        ~((base['PARTIDAOF'].isin([39908, 39909])) & (base['FUENTE_FIN'] != 'Rec. Propios') & (base['UNIDAD_OP_NOM'] != 'OFICINAS CENTRALES  '))
    ]

    # Gasto transversal
    gasto_trans = gasto[
        (((gasto['PROG_PRES'] == 'M001') & (gasto['PARTIDAOF'] != 33903)) |
        (gasto['PROG_PRES'].isin(['O001', 'P021']))) &
        (~gasto['PARTIDAOF'].isin([39908, 39909]))
    ]
    # Optimización: En lugar de groupby completo, calculamos directo por capítulos necesarios
    capitulos_interes = [10000000, 20000000, 30000000, 40000000, 50000000]
    
    g_trans_dict = (gasto_trans.groupby('CAPITULO')['IMPORTE'].sum() / 1000000).to_dict()
    ta = gasto_trans['IMPORTE'].sum() / 1000000

    # Gasto Acopio
    gasto_acopiop = gasto[(gasto['PROG_PRES'] == 'S290') & (gasto['FUENTE_FIN'] == 'Rec. Propios') & (~gasto['PARTIDAOF'].isin([39908, 39909]))]
    gasto_acopiof = gasto[(gasto['PROG_PRES'] == 'S290') & (gasto['FUENTE_FIN'] != 'Rec. Propios')]
    
    gap_dict = (gasto_acopiop.groupby('CAPITULO')['IMPORTE'].sum() / 1000000).to_dict()
    gaf = gasto_acopiof['IMPORTE'].sum() / 1000000
    a = sum(gap_dict.values())

    # Gasto Par
    excluir_areas = [952, 953, 954, 955, 957]
    gasto_par_base = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Propios') & (~gasto['AREA_AFECT'].isin(excluir_areas)) & (~gasto['PARTIDAOF'].isin([39908, 39909]))]
    gasto_p = gasto[(gasto['PROG_PRES'] == 'M001') & (gasto['PARTIDAOF'] == 33903)]
    
    gpp_dict = gasto_par_base.groupby('CAPITULO')['IMPORTE'].sum().to_dict()
    gpp2_dict = gasto_p.groupby('CAPITULO')['IMPORTE'].sum().to_dict()
    
    # Combinar diccionarios de gasto par propio
    gpp4_dict = {}
    for c in capitulos_interes:
        gpp4_dict[c] = gpp_dict.get(c, 0) + gpp2_dict.get(c, 0)
    p = sum(gpp4_dict.values()) / 1000000

    gpf = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & (~gasto['CAPITULO'].isin([40000000, 10000000])) & (~gasto['AREA_AFECT'].isin(excluir_areas))]
    gpf2_sum = gasto[gasto['ID'].isin(gpf['ID'])]['IMPORTE'].sum()
    pf = gpf2_sum / 1000000

    # Gasto Transformación
    gasto_tp1 = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Propios') & (gasto['AREA_AFECT'].isin(excluir_areas)) & (~gasto['PARTIDAOF'].isin([39908, 39909]))]
    gasto_tp2 = gasto[(gasto['PROG_PRES'] == 'W001') & (gasto['PARTIDA'].isin([39909003, 39909005])) & (gasto['AREA_AFECT'].isin(excluir_areas))]
    
    gtp_dict = (pd.concat([gasto_tp1, gasto_tp2]).groupby('CAPITULO')['IMPORTE'].sum() / 1000000).to_dict()
    tr = sum(gtp_dict.values())

    gasto_tf = gasto[((gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & (gasto['CAPITULO'] == 40000000)) | (((gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & (gasto['CAPITULO'] != 40000000)) & (gasto['AREA_AFECT'].isin(excluir_areas)))]
    trf = gasto_tf['IMPORTE'].sum() / 1000000

    # Gasto Maíz
    gmr1 = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & (gasto['CAPITULO'] == 10000000)]
    gmr = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['AREA_AFECT'] == 990) & (gasto['FUENTE_FIN'] == 'Rec. Fiscales')]
    gmaiz_dict = (pd.concat([gmr1, gmr]).groupby('CAPITULO')['IMPORTE'].sum() / 1000000).to_dict()
    m = sum(gmaiz_dict.values())

    # --- CONSTRUCCIÓN DE MATRIZ DE EGRESOS ---
    conceptos_gastos = {
        10000000: 'Servicios Personales',
        20000000: 'Materiales y Suministros',
        30000000: 'Servicios Generales',
        40000000: 'Subsidios y Transferencias',
        50000000: 'Inversión'
    }

    filas_egresos = []
    for cap, concepto in conceptos_gastos.items():
        # Lógica de asignación exacta para Acopio (Cap 40M une fiscales)
        v_acopio = gap_dict.get(cap, 0)
        if cap == 40000000:
            v_acopio += gaf
            
        # Lógica PAR (Cap 40M une fiscales)
        v_par = gpp4_dict.get(cap, 0) / 1000000
        if cap == 40000000:
            v_par += pf
            
        # Lógica Transformación
        v_transf = gtp_dict.get(cap, 0)
        if cap == 40000000:
            v_transf += trf

        filas_egresos.append({
            'Concepto': concepto,
            'Acopio': f"{v_acopio:,.3f}",
            'PAR': f"{v_par:,.3f}",
            'Transformación': f"{v_transf:,.3f}",
            'Maíz es la Raíz': f"{gmaiz_dict.get(cap, 0):,.3f}",
            'Gasto Transversal': f"{g_trans_dict.get(cap, 0):,.3f}"
        })
    
    df_egresos = pd.DataFrame(filas_egresos)

    # --- SECCIÓN INGRESOS ---
    pf_ing = base[(base['PARTIDAOF'] == 72310) & (base['FUENTE_FIN'] == 'Rec. Propios')]['IMPORTE'].sum() / 1000000
    otros_ing = base[(base['PARTIDAOF'] == 72320) & (base['FUENTE_FIN'] == 'Rec. Propios')]['IMPORTE'].sum() / 1000000
    ventas_sum = abs(base[(base['PARTIDAOF'] == 72210) & (base['FUENTE_FIN'] == 'Rec. Propios')]['IMPORTE'].sum() / 1000000)

    df_ingresos = pd.DataFrame([
        {'Concepto': 'Venta de Bienes', 'Acopio': f"{494.370:,.3f}", 'PAR': f"{(ventas_sum - 494.37):,.3f}", 'Transformación': "0.000", 'Maíz es la Raíz': "0.000", 'Gasto Transversal': "0.000"},
        {'Concepto': 'Productos Financieros', 'Acopio': "0.000", 'PAR': "0.000", 'Transformación': "0.000", 'Maíz es la Raíz': "0.000", 'Gasto Transversal': f"{abs(pf_ing):,.3f}"},
        {'Concepto': 'Otros', 'Acopio': "0.000", 'PAR': f"{abs(otros_ing):,.3f}", 'Transformación': "0.000", 'Maíz es la Raíz': "0.000", 'Gasto Transversal': "0.000"},
        {'Concepto': 'Subsidios y Transferencias', 'Acopio': "0.000", 'PAR': "0.000", 'Transformación': "0.000", 'Maíz es la Raíz': "0.000", 'Gasto Transversal': "0.000"}
    ])

    # Secciones divisoras limpias
    df_div_ingresos = pd.DataFrame([{'Concepto': 'Ingresos', 'Acopio': '', 'PAR': '', 'Transformación': '', 'Maíz es la Raíz': '', 'Gasto Transversal': ''}])
    df_div_egresos = pd.DataFrame([{'Concepto': 'Egresos', 'Acopio': '', 'PAR': '', 'Transformación': '', 'Maíz es la Raíz': '', 'Gasto Transversal': ''}])

    # --- FUENTES DE FINANCIAMIENTO ---
    af = gaf # Rescate del valor anterior
    df_fuente = pd.DataFrame([
        {'Concepto': 'Fuente de Financiamiento', 'Acopio': '', 'PAR': '', 'Transformación': '', 'Maíz es la Raíz': '', 'Gasto Transversal': ''},
        {'Concepto': 'Propios', 'Acopio': f"{a:,.3f}", 'PAR': f"{p:,.3f}", 'Transformación': f"{tr:,.3f}", 'Maíz es la Raíz': "0.000", 'Gasto Transversal': f"{ta:,.3f}"},
        {'Concepto': 'Fiscales', 'Acopio': f"{af:,.3f}", 'PAR': f"{pf:,.3f}", 'Transformación': f"{trf:,.3f}", 'Maíz es la Raíz': f"{m:,.3f}", 'Gasto Transversal': "0.000"}
    ])

    # Ensamble final estructurado sin colapsos de índices recurrentes
    tabla_final = pd.concat([df_div_ingresos, df_ingresos, df_div_egresos, df_egresos, df_fuente], ignore_index=True)
    return tabla_final 

# 3. Interfaz de usuario
archivo = st.file_uploader("Sube tu archivo de Excel", type=["xlsx"])
boton = st.button('Genera el BFP')

# 4. Ejecución de la lógica
if archivo is not None:
    try:
        st.session_state.df_original = cargar_datos(archivo)
        df_original = st.session_state.df_original
        st.success("¡Archivo cargado con éxito en memoria!")
        st.write("Vista previa de los datos:")
        st.dataframe(df_original.head(5), width="stretch")
    except Exception as e:
        st.error(f"Hubo un error al procesar el archivo: {e}")
        st.stop()

if boton is True:
    df = st.dataframe(st.session_state.df_original)
    tabla = balance(df)
    st.write(tabla, width="stretch")


    

