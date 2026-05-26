import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Configuración de la página
st.set_page_config(page_title="Generador de Balance", page_icon="📊", layout="wide")

st.title("📊 Generador de Reporte de Balance")
st.markdown("Sube tu archivo de Excel para procesar los ingresos, egresos y fuentes de financiamiento.")

uploaded_file = st.file_uploader("Elige un archivo de Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("Procesando datos de manera optimizada..."):
        try:
            # 1. Cargar datos eficientemente
            datos = pd.read_excel(uploaded_file, sheet_name=0)
            df = pd.DataFrame(datos)

            # 2. Limpieza básica
            df['FECHA'] = pd.to_datetime(df['FECHA_CONEX'], format='%d/%m/%Y', errors='coerce')
            df_limpio1 = df.dropna(subset=['FECHA'])
            df_limpio2 = df_limpio1[df_limpio1['ESTATUS'] == 'Conectado']
            df_limpio3 = df_limpio2[df_limpio2['TIPO_CEGAP'] != 'Cegap Nac. de Proveedores Descontados']

            # Filtro base
            cond_duplicado = (
                (df_limpio3['FOLIO_DEF'].between(49999, 79999)) &
                (df_limpio3['TIPO_CEGAP'].isin(['Cegap de Erogacion', 'Cegap de Erogacion de Proveedores (Mercancias)'])) &
                (df_limpio3['TIPO_PAGO'] == 'Cegap de registro')
            )
            base = df_limpio3[~cond_duplicado].copy()

            # --- SECCIÓN: GASTO (EGRESOS) ---
            gasto = base[
                (base['CAPITULO'] != 7000000) &
                (base['PARTIDA'] != 43101001) &
                (base['PARTIDA'] != 43701001) &
                (base['PARTIDA'] != 39908031) &
                (base['PARTIDA'] != 39908046)
            ].copy()
            
            # Filtro compuesto para PARTIDAOF
            cond_partidaof = (gasto['PARTIDAOF'].isin([39908, 39909]))
            gasto = gasto[~(cond_partidaof & (gasto['FUENTE_FIN'] != 'Rec. Propios') & (gasto['UNIDAD_OP_NOM'] != 'OFICINAS CENTRALES  '))]

            # Gasto Transversal
            gasto_trans = gasto[
                (((gasto['PROG_PRES'] == 'M001') & (gasto['PARTIDAOF'] != 33903)) |
                 (gasto['PROG_PRES'].isin(['O001', 'P021']))) &
                ~cond_partidaof
            ]
            gasto_transversal = (gasto_trans.groupby('CAPITULO')['IMPORTE'].sum() / 1000000).reset_index()
            gasto_transversal = gasto_transversal.rename(columns={'CAPITULO': 'Capítulo', 'IMPORTE': 'Gasto Transversal'})

            # Gasto Acopio
            gasto_acopiop = gasto[(gasto['PROG_PRES'] == 'S290') & (gasto['FUENTE_FIN'] == 'Rec. Propios') & ~cond_partidaof]
            gasto_acopiof = gasto[(gasto['PROG_PRES'] == 'S290') & (gasto['FUENTE_FIN'] != 'Rec. Propios')]
            
            gap = gasto_acopiop.groupby('CAPITULO')['IMPORTE'].sum() / 1000000
            gaf = gasto_acopiof['IMPORTE'].sum() / 1000000
            
            ga = pd.concat([gap, pd.Series([gaf], index=[40000000])], axis=0)
            gasto_acopio = ga.reset_index().rename(columns={'index': 'Capítulo', 0: 'Acopio'})

            # Gasto PAR (Optimizado sin usar el .isin() masivo)
            areas_excluidas = [952, 953, 954, 955, 957]
            gasto_par_df = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Propios') & ~(gasto['AREA_AFECT'].isin(areas_excluidas)) & ~cond_partidaof]
            gpp = gasto_par_df.groupby('CAPITULO')['IMPORTE'].sum()

            gasto_p = gasto[(gasto['PROG_PRES'] == 'M001') & (gasto['PARTIDAOF'] == 33903)]
            gpp2 = gasto_p.groupby('CAPITULO')['IMPORTE'].sum()
            gpp4 = pd.concat([gpp, gpp2], axis=0).groupby(level=0).sum()

            # OPTIMIZACIÓN AQUÍ: En lugar de buscar IDs con .isin(), calculamos directamente sobre el filtro original recortado
            gpf_subset = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & ~(gasto['CAPITULO'].isin([40000000, 10000000])) & ~(gasto['AREA_AFECT'].isin(areas_excluidas))]
            gpf2_val = gpf_subset['IMPORTE'].sum()
            
            gpar = pd.concat([gpp4, pd.Series([gpf2_val], index=[40000000])], axis=0).groupby(level=0).sum()
            gasto_par_final = (gpar / 1000000).reset_index().rename(columns={'index': 'Capítulo', 'IMPORTE': 'PAR', 0: 'PAR'})

            # Gasto Transformación
            gasto_tp1 = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Propios') & (gasto['AREA_AFECT'].isin(areas_excluidas)) & ~cond_partidaof]
            gasto_tp2 = gasto[(gasto['PROG_PRES'] == 'W001') & (gasto['PARTIDA'].isin([39909003, 39909005])) & (gasto['AREA_AFECT'].isin(areas_excluidas))]
            g_tp = pd.concat([gasto_tp1, gasto_tp2], axis=0).groupby('CAPITULO')['IMPORTE'].sum() / 1000000

            gasto_tf = gasto[((gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & (gasto['CAPITULO'] == 40000000)) |
                             ((gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & (gasto['CAPITULO'] != 40000000) & (gasto['AREA_AFECT'].isin(areas_excluidas)))]
            g_tf = gasto_tf['IMPORTE'].sum() / 1000000
            
            g_tpf = pd.concat([g_tp, pd.Series([g_tf], index=[40000000])], axis=0).groupby(level=0).sum()
            gasto_transformacion = g_tpf.reset_index().rename(columns={'index': 'Capítulo', 0: 'Transformación'})

            # Gasto Maíz es la Raíz
            gmr1 = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & (gasto['CAPITULO'] == 10000000)]
            gastom = gasto1 = gmr1.groupby('CAPITULO')['IMPORTE'].sum() / 1000000
            gmr = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['AREA_AFECT'] == 990) & (gasto['FUENTE_FIN'] == 'Rec. Fiscales')]
            gastom2 = gmr.groupby('CAPITULO')['IMPORTE'].sum() / 1000000
            gasto_maiz = pd.concat([gastom, gastom2], axis=0).groupby(level=0).sum().reset_index().rename(columns={'index': 'Capítulo', 0: 'Maíz es la Raíz'})

            # Unir Egresos
            gasto_cap = pd.DataFrame({
                'Capítulo': [10000000, 20000000, 30000000, 40000000, 50000000],
                'Concepto': ['Servicios Personales', 'Materiales y Suministros', 'Servicios Generales', 'Subsidios y Transferencias', 'Inversión']
            })

            e1 = pd.merge(gasto_cap, gasto_acopio, on='Capítulo', how='outer')
            e2 = pd.merge(e1, gasto_par_final, on='Capítulo', how='outer')
            e3 = pd.merge(e2, gasto_transformacion, on='Capítulo', how='outer')
            e4 = pd.merge(e3, gasto_maiz, on='Capítulo', how='outer')
            e5 = pd.merge(e4, gasto_transversal, on='Capítulo', how='outer').fillna(0)
            
            # Guardamos numéricos para fuentes de financiamiento
            a_val = gap.sum()
            p_val = gpp4.sum() / 1000000
            tr_val = g_tp.sum()
            ta_val = gasto_transversal['Gasto Transversal'].sum() if not gasto_transversal.empty else 0
            af_val = gaf
            pf_val = gpf2_val / 1000000
            trf_val = g_tf
            m_val = gasto_maiz['Maíz es la Raíz'].sum() if not gasto_maiz.empty else 0

            # Formatear strings
            egresos = e5.copy()
            for col in ['Acopio', 'PAR', 'Transformación', 'Maíz es la Raíz', 'Gasto Transversal']:
                egresos[col] = egresos[col].map('{:,.3f}'.format)
            egeresos2 = egresos.drop('Capítulo', axis=1)

            # --- SECCIÓN: INGRESO ---
            pf = base[(base['PARTIDAOF'] == 72310) & (base['FUENTE_FIN'] == 'Rec. Propios')]
            pf2_val = (pf['IMPORTE'].sum() / 1000000)
            productos_financieros = pd.DataFrame({'Concepto': ['Productos Financieros'], 'Gasto Transversal': [abs(pf2_val)], 'PAR': [0], 'Acopio': [0], 'Transformación': [0], 'Maíz es la Raíz': [0]})

            otros = base[(base['PARTIDAOF'] == 72320) & (base['FUENTE_FIN'] == 'Rec. Propios')]
            otros2_val = (otros['IMPORTE'].sum() / 1000000)
            otros_ingresos = pd.DataFrame({'Concepto': ['Otros'], 'PAR': [abs(otros2_val)], 'Acopio': [0], 'Transformación': [0], 'Maíz es la Raíz': [0], 'Gasto Transversal': [0]})

            ventas = base[(base['PARTIDAOF'] == 72210) & (base['FUENTE_FIN'] == 'Rec. Propios')]
            ventas2 = abs((ventas['IMPORTE'].sum() / 1000000))
            ventas_bienes = pd.DataFrame({'Concepto': ['Venta de Bienes'], 'Acopio': [494.37], 'PAR': [ventas2 - 494.37], 'Transformación': [0], 'Maíz es la Raíz': [0], 'Gasto Transversal': [0]})

            ingresos = pd.DataFrame({
                'Concepto': ['Venta de Bienes', 'Productos Financieros', 'Otros', 'Subsidios y Transferencias'],
                'Acopio': [ventas_bienes['Acopio'].sum(), productos_financieros['Acopio'].sum(), otros_ingresos['Acopio'].sum(), 0],
                'PAR': [ventas_bienes['PAR'].sum(), productos_financieros['PAR'].sum(), otros_ingresos['PAR'].sum(), 0],
                'Transformación': [0, 0, 0, 0],
                'Maíz es la Raíz': [0, 0, 0, 0],
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

            tabla = pd.concat([resultado, fuente], axis=0).reset_index().rename(columns={'index': 'Concepto'})

            st.success("¡Reporte generado exitosamente!")
            st.dataframe(tabla, use_container_width=True)

            # Botón de Descarga
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                tabla.to_excel(writer, index=False, sheet_name='Balance')
            st.download_button(
                label="📥 Descargar Balance en Excel",
                data=output.getvalue(),
                file_name="tabla_balance.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error en el procesamiento: {e}")
