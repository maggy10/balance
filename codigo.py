import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Configuración de la página
st.set_page_config(page_title="Generador de Balance", page_icon="📊", layout="wide")

st.title("📊 Generador de Reporte de Balance")
st.markdown("Sube tu archivo de Excel para procesar los ingresos, egresos y fuentes de financiamiento.")

# Componente para subir el archivo
uploaded_file = st.file_uploader("Elige un archivo de Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("Procesando datos... Por favor espera."):
        try:
            # Cargar datos (Leemos la primera pestaña tal como en tu código original)
            datos = pd.read_excel(uploaded_file, sheet_name=0)
            df = pd.DataFrame(datos)

            # --- LIMPIEZA Y PREPARACIÓN ---
            df['FECHA'] = pd.to_datetime(df['FECHA_CONEX'], format='%d/%m/%Y', errors='coerce')
            df['ID'] = df['SUCURSAL_CVE'].astype(str) + df['UNIDAD_OP_CVE'].astype(str) + df['FOLIO_DEF'].astype(str)
            
            df_limpio1 = df.dropna(subset=['FECHA'])
            df_limpio2 = df_limpio1[df_limpio1['ESTATUS'] == 'Conectado']
            df_limpio3 = df_limpio2[df_limpio2['TIPO_CEGAP'] != 'Cegap Nac. de Proveedores Descontados']

            base = df_limpio3[~((df_limpio3['FOLIO_DEF'].between(49999, 79999)) &
                                ((df_limpio3['TIPO_CEGAP'] == 'Cegap de Erogacion') |
                                (df_limpio3['TIPO_CEGAP'] == 'Cegap de Erogacion de Proveedores (Mercancias)')) &
                                (df_limpio3['TIPO_PAGO'] == 'Cegap de registro'))]

            # --- SECCIÓN: GASTO (EGRESOS) ---
            gasto = base[(base['CAPITULO'] != 7000000)]
            gasto = gasto[(gasto['PARTIDA'] != 43101001)]
            gasto = gasto[(gasto['PARTIDA'] != 43701001)]
            gasto = gasto[(gasto['PARTIDA'] != 39908031)]
            gasto = gasto[(gasto['PARTIDA'] != 39908046)]
            gasto = gasto[~(((gasto['PARTIDAOF'] == 39908) | (gasto['PARTIDAOF'] == 39909)) &
                            (gasto['FUENTE_FIN'] != 'Rec. Propios') &
                            (gasto['UNIDAD_OP_NOM'] != 'OFICINAS CENTRALES  '))]

            # Gasto Transversal
            gasto_trans = gasto[(((gasto['PROG_PRES'] == 'M001') & (gasto['PARTIDAOF'] != 33903)) |
                                 (gasto['PROG_PRES'] == 'O001') |
                                 (gasto['PROG_PRES'] == 'P021')) &
                                ~(((gasto['PARTIDAOF'] == 39908) | (gasto['PARTIDAOF'] == 39909)))]
            gasto_transversal = ((gasto_trans.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).reset_index()
            gasto_transversal = gasto_transversal.rename(columns={'CAPITULO': 'Capítulo', 'IMPORTE': 'Gasto Transversal'})

            # Gasto Acopio
            gasto_acopiop = gasto[(gasto['PROG_PRES'] == 'S290') & 
                                  (gasto['FUENTE_FIN'] == 'Rec. Propios') &
                                  ~(((gasto['PARTIDAOF'] == 39908) | (gasto['PARTIDAOF'] == 39909)))]
            gasto_acopiof = gasto[(gasto['PROG_PRES'] == 'S290') & (gasto['FUENTE_FIN'] != 'Rec. Propios')]
            
            gap = ((gasto_acopiop.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).reset_index().set_index('CAPITULO')
            gaf = ((gasto_acopiof['IMPORTE'].sum()) / 1000000)
            gaf2 = pd.DataFrame({'CAPITULO': [40000000], 'IMPORTE': [gaf]}).set_index('CAPITULO')
            ga = pd.concat([gap, gaf2], axis=0)
            gasto_acopio = ga.reset_index().rename(columns={'CAPITULO': 'Capítulo', 'IMPORTE': 'Acopio'})

            # Gasto PAR
            gasto_par_df = gasto[((gasto['PROG_PRES'] == 'S053') & 
                                  (gasto['FUENTE_FIN'] == 'Rec. Propios') &
                                  ~((gasto['AREA_AFECT'].isin([952, 953, 954, 955, 957])))) &
                                 ~((gasto['PARTIDAOF'].isin([39908, 39909])))]
            gpp = gasto_par_df.groupby('CAPITULO')['IMPORTE'].sum()

            gasto_p = gasto[(gasto['PROG_PRES'] == 'M001') & (gasto['PARTIDAOF'] == 33903)]
            gpp2 = gasto_p.groupby('CAPITULO')['IMPORTE'].sum()
            gpp3 = (pd.concat([gpp, gpp2], axis=0)).reset_index()
            gpp4 = (gpp3.groupby('CAPITULO')['IMPORTE'].sum()).reset_index()

            gpf = gasto[((gasto['PROG_PRES'] == 'S053') & 
                         (gasto['FUENTE_FIN'] == 'Rec. Fiscales') &
                         (~gasto['CAPITULO'].isin([40000000, 10000000]))) &
                        ~((gasto['AREA_AFECT'].isin([952, 953, 954, 955, 957])))]
            gpf2 = gasto[gasto['ID'].isin(gpf['ID'])]['IMPORTE'].sum()
            gpf3 = pd.DataFrame({'CAPITULO': [40000000], 'IMPORTE': [gpf2]})
            gpf3.index.name = 'CAPITULO'

            gpar = pd.concat([gpp4, gpf3], axis=0)
            gasto_par_final = ((gpar.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).reset_index()
            gasto_par_final = gasto_par_final.rename(columns={'CAPITULO': 'Capítulo', 'IMPORTE': 'PAR'})

            # Gasto Transformación
            gasto_tp1 = gasto[((gasto['PROG_PRES'] == 'S053') & 
                               (gasto['FUENTE_FIN'] == 'Rec. Propios') &
                               (gasto['AREA_AFECT'].isin([952, 953, 954, 955, 957]))) &
                              ~gasto['PARTIDAOF'].isin([39908, 39909])]
            gasto_tp2 = gasto[(gasto['PROG_PRES'] == 'W001') & 
                              (gasto['PARTIDA'].isin([39909003, 39909005])) &
                              (gasto['AREA_AFECT'].isin([952, 953, 954, 955, 957]))]
            gasto_tp = pd.concat([gasto_tp1, gasto_tp2], axis=0)
            g_tp = ((gasto_tp.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).reset_index()

            gasto_tf = gasto[((gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & (gasto['CAPITULO'] == 40000000)) |
                             (((gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & (gasto['CAPITULO'] != 40000000)) &
                              (gasto['AREA_AFECT'].isin([952, 953, 954, 955, 957])))]
            g_tf = (gasto_tf['IMPORTE'].sum()) / 1000000
            g_tf2 = pd.DataFrame({'CAPITULO': [40000000], 'IMPORTE': [g_tf]})
            g_tf2.index.name = 'CAPITULO'

            g_tpf = pd.concat([g_tp, g_tf2], axis=0)
            gasto_transformacion = ((g_tpf.groupby('CAPITULO')['IMPORTE'].sum())).reset_index()
            gasto_transformacion = gasto_transformacion.rename(columns={'CAPITULO': 'Capítulo', 'IMPORTE': 'Transformación'})

            # Gasto Maíz es la Raíz
            gmr1 = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & (gasto['CAPITULO'] == 10000000)]
            gastom = ((gmr1.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).reset_index()
            gmr = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['AREA_AFECT'] == 990) & (gasto['FUENTE_FIN'] == 'Rec. Fiscales')]
            gastom2 = ((gmr.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).reset_index()
            g_tpf_maiz = pd.concat([gastom, gastom2], axis=0)
            gasto_maiz = g_tpf_maiz.rename(columns={'CAPITULO': 'Capítulo', 'IMPORTE': 'Maíz es la Raíz'})

            # Unir Egresos
            gasto_cap = pd.DataFrame({
                'Capítulo': [10000000, 20000000, 30000000, 40000000, 50000000],
                'Concepto': ['Servicios Personales', 'Materiales y Suministros', 'Servicios Generales', 'Subsidios y Transferencias', 'Inversión']
            }).set_index('Capítulo')

            e1 = pd.merge(gasto_cap, gasto_acopio, on='Capítulo', how='outer')
            e2 = pd.merge(e1, gasto_par_final, on='Capítulo', how='outer')
            e3 = pd.merge(e2, gasto_transformacion, on='Capítulo', how='outer')
            e4 = pd.merge(e3, gasto_maiz, on='Capítulo', how='outer')
            e5 = pd.merge(e4, gasto_transversal, on='Capítulo', how='outer')
            
            egresos = e5.fillna(0)
            
            # Guardamos numéricos antes de dar formato de texto para usarlos en "fuente"
            a_val = gap['IMPORTE'].sum() if not gap.empty else 0
            p_val = gpp4['IMPORTE'].sum() / 1000000 if not gpp4.empty else 0
            tr_val = g_tp['IMPORTE'].sum() if not g_tp.empty else 0
            ta_val = gasto_transversal['Gasto Transversal'].sum() if not gasto_transversal.empty else 0
            
            af_val = gaf2['IMPORTE'].sum() if not gaf2.empty else 0
            pf_val = gpf3['IMPORTE'].sum() / 1000000 if not gpf3.empty else 0
            trf_val = g_tf2['IMPORTE'].sum() if not g_tf2.empty else 0
            m_val = gasto_maiz['Maíz es la Raíz'].sum() if not gasto_maiz.empty else 0

            # Formatear strings para mostrar en la app
            for col in ['Acopio', 'PAR', 'Transformación', 'Maíz es la Raíz', 'Gasto Transversal']:
                egresos[col] = egresos[col].map('{:,.3f}'.format)
            egeresos2 = egresos.drop('Capítulo', axis=1)

            # --- SECCIÓN: INGRESO ---
            pf = base[(base['PARTIDAOF'] == 72310) & (base['FUENTE_FIN'] == 'Rec. Propios')]
            pf2 = ((pf.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).abs().reset_index()
            pf2_val = pf2['IMPORTE'].sum() if not pf2.empty else 0
            productos_financieros = pd.DataFrame({'Concepto': ['Productos Financieros'], 'Gasto Transversal': [pf2_val], 'PAR': [0], 'Acopio': [0], 'Transformación': [0], 'Maíz es la Raíz': [0]})

            otros = base[(base['PARTIDAOF'] == 72320) & (base['FUENTE_FIN'] == 'Rec. Propios')]
            otros2 = ((otros.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).abs().reset_index()
            otros2_val = otros2['IMPORTE'].sum() if not otros2.empty else 0
            otros_ingresos = pd.DataFrame({'Concepto': ['Otros'], 'PAR': [otros2_val], 'Acopio': [0], 'Transformación': [0], 'Maíz es la Raíz': [0], 'Gasto Transversal': [0]})

            ventas = base[(base['PARTIDAOF'] == 72210) & (base['FUENTE_FIN'] == 'Rec. Propios')]
            ventas2 = abs(((ventas['IMPORTE'].sum()) / 1000000))
            ventas_bienes = pd.DataFrame({'Concepto': ['Venta de Bienes'], 'Acopio': [494.37], 'PAR': [ventas2 - 494.37], 'Transformación': [0], 'Maíz es la Raíz': [0], 'Gasto Transversal': [0]})

            ingresos = pd.DataFrame({
                'Concepto': ['Venta de Bienes', 'Productos Financieros', 'Otros', 'Subsidios y Transferencias'],
                'Acopio': [ventas_bienes['Acopio'].sum(), productos_financieros['Acopio'].sum(), otros_ingresos['Acopio'].sum(), 0],
                'PAR': [ventas_bienes['PAR'].sum(), productos_financieros['PAR'].sum(), otros_ingresos['PAR'].sum(), 0],
                'Transformación': [ventas_bienes['Transformación'].sum(), productos_financieros['Transformación'].sum(), otros_ingresos['Transformación'].sum(), 0],
                'Maíz es la Raíz': [ventas_bienes['Maíz es la Raíz'].sum(), productos_financieros['Maíz es la Raíz'].sum(), otros_ingresos['Maíz es la Raíz'].sum(), 0],
                'Gasto Transversal': [ventas_bienes['Gasto Transversal'].sum(), productos_financieros['Gasto Transversal'].sum(), otros_ingresos['Gasto Transversal'].sum(), 0]
            })

            for col in ['Acopio', 'PAR', 'Transformación', 'Maíz es la Raíz', 'Gasto Transversal']:
                ingresos[col] = ingresos[col].map('{:,.3f}'.format)

            # --- CONCATENAR RESULTADOS ---
            ig = pd.DataFrame({'Concepto': ['Acopio', 'PAR', 'Transformación', 'Maíz es la Raíz', 'Gasto Transversal'], 'Ingresos': ['', '', '', '', '']}).set_index('Concepto').T
            eg = pd.DataFrame({'Concepto': ['Acopio', 'PAR', 'Transformación', 'Maíz es la Raíz', 'Gasto Transversal'], 'Egresos': ['', '', '', '', '']}).set_index('Concepto').T

            ingresos = ingresos.set_index('Concepto')
            egeresos2 = egeresos2.set_index('Concepto')

            resultado = pd.concat([ig, ingresos, eg, egeresos2], axis=0)

            # --- FUENTE DE FINANCIAMIENTO ---
            fuente = pd.DataFrame({
                'Concepto': ['Acopio', 'PAR', 'Transformación', 'Maíz es la Raíz', 'Gasto Transversal'],
                'Fuente de Financiamiento': ['', '', '', '', ''],
                'Propios': [a_val, p_val, tr_val, 0, ta_val],
                'Fiscales': [af_val, pf_val, trf_val, m_val, 0]
            }).set_index('Concepto')

            fuente['Propios'] = fuente['Propios'].map('{:,.3f}'.format)
            fuente['Fiscales'] = fuente['Fiscales'].map('{:,.3f}'.format)
            fuente = fuente.T

            # TABLA FINAL
            tabla = pd.concat([resultado, fuente], axis=0).reset_index().rename(columns={'index': 'Concepto'})

            # --- MOSTRAR EN STREAMLIT ---
            st.success("¡Archivo procesado con éxito!")
            
            st.subheader("Vista Previa del Balance")
            st.dataframe(tabla, use_container_width=True)

            # --- DESCARGAR EXCEL ---
            # Guardamos el Excel en un buffer de memoria para que el usuario pueda descargarlo directamente
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                tabla.to_excel(writer, index=False, sheet_name='Balance')
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Descargar Tabla en Excel",
                data=processed_data,
                file_name="tabla_balance.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Ocurrió un error al procesar el archivo: {e}")
