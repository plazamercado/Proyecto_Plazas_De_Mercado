import pandas as pd
import os
import sys
import time

# --- CONFIGURACIÓN ---
# ID de tu hoja de cálculo de Google Sheets (DEBE SER PÚBLICO)
SHEET_ID = '1rryqAVgJBGBTNj_fGLfD4Gg_WWQjfPSe5WNWKGoGoYg' 
SHEET_NAME = 'Hoja1' # Asegúrate que coincida con el nombre de tu pestaña
URL = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}&cache_buster={int(time.time())}'

def generar_reportes_completos():
    try:
        print("Descargando datos...")
        df = pd.read_csv(URL)

        # --- LIMPIEZA GENERAL ---
        df = df.fillna(0)
        df['VENTA_PRECIO'] = pd.to_numeric(df['VENTA_PRECIO'], errors='coerce').fillna(0)

        if 'TIPO_PUNTO' in df.columns:
            df['TIPO_PUNTO'] = df['TIPO_PUNTO'].apply(
                lambda x: 'plaza' if 'plaza' in str(x).lower() or 'pmd' in str(x).lower() else 'externo'
            )

        if 'FECHA' in df.columns:
            # dayfirst=True soluciona el Warning de formato de fecha
            df['FECHA'] = pd.to_datetime(df['FECHA'], dayfirst=True, errors='coerce')
            df['FECHA_DIA'] = df['FECHA'].dt.date
            df = df.dropna(subset=['FECHA'])

        df_limpio = df[df['VENTA_PRECIO'] > 0].copy()

        output_file = 'Reporte_Comparativo.xlsx'

        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Formatos
            num_fmt = workbook.add_format({'num_format': '#,##0'})
            percent_fmt = workbook.add_format({'num_format': '0.00%'})
            total_text_fmt = workbook.add_format({'bg_color': '#E7E6E6', 'bold': True})
            total_num_fmt = workbook.add_format({'bg_color': '#E7E6E6', 'bold': True, 'num_format': '#,##0'})

            # =====================================================
            # HOJA 1: RESUMEN PMD
            # =====================================================
            resumen = df_limpio.pivot_table(
                index='PLAZA', 
                columns='TIPO_PUNTO', 
                values='VENTA_PRECIO', 
                aggfunc='sum', 
                fill_value=0
            ).reset_index()

            resumen = resumen.rename(columns={'PLAZA': 'PDM', 'plaza': '$PM', 'externo': '$Tienda'})
            
            # Asegurar que existan las columnas para cálculos
            for col in ['$PM', '$Tienda']:
                if col not in resumen.columns: resumen[col] = 0

            resumen['Diferencia (Tienda - Plaza)'] = resumen['$Tienda'] - resumen['$PM']
            resumen['Represent %'] = resumen.apply(
                lambda r: (r['Diferencia (Tienda - Plaza)'] / r['$PM']) if r['$PM'] != 0 else 0, axis=1
            )
            
            resumen = resumen[['PDM', '$PM', '$Tienda', 'Diferencia (Tienda - Plaza)', 'Represent %']]
            
            # Crear espacio de 2 filas y promedio sin Warnings
            vacias = pd.DataFrame([{c: None for c in resumen.columns}] * 2)
            promedio = pd.DataFrame([{
                'PDM': 'Promedio', 
                '$PM': resumen['$PM'].sum(),
                '$Tienda': None, 'Diferencia (Tienda - Plaza)': None, 'Represent %': None
            }])
            
            resumen_final = pd.concat([resumen, vacias, promedio], ignore_index=True)
            resumen_final.to_excel(writer, sheet_name='Resumen PMD', index=False)
            
            ws_res = writer.sheets['Resumen PMD']
            ws_res.set_column('B:D', 18, num_fmt)
            ws_res.set_column('E:E', 15, percent_fmt)

            # =====================================================
            # HOJA 2: PRECIOS SDDE
            # =====================================================
            precios_sdde = df_limpio.pivot_table(
                index='FECHA_DIA', 
                columns='PRODUCTO', 
                values='VENTA_PRECIO', 
                aggfunc='mean'
            ).round(0)
            
            precios_sdde.to_excel(writer, sheet_name='Precios SDDE')
            ws_sdde = writer.sheets['Precios SDDE']
            ws_sdde.set_column('A:A', 20)
            ws_sdde.set_column('B:XFD', 12, num_fmt)
            ws_sdde.freeze_panes(1, 1)

            # =====================================================
            # HOJAS POR PLAZA
            # =====================================================
            plazas = df_limpio['PLAZA'].dropna().unique()
            
            for plaza in plazas:
                df_pla = df_limpio[
                    (df_limpio['PLAZA'] == plaza) & 
                    (df_limpio['ES_CANASTA'].str.upper().isin(['SI', 'SÍ']))
                ]
                if df_pla.empty: continue
                
                reporte = df_pla.pivot_table(
                    index=['GRUPO_ALIMENTARIO', 'PRODUCTO'], 
                    columns='TIPO_PUNTO', 
                    values='VENTA_PRECIO', 
                    aggfunc='mean', 
                    fill_value=0
                ).reset_index()

                # Garantizar columnas de plaza y externo
                if 'plaza' not in reporte.columns: reporte['plaza'] = 0
                if 'externo' not in reporte.columns: reporte['externo'] = 0

                nombre_pdm = f"PDM {plaza}" if "PDM" not in str(plaza).upper() else str(plaza)
                reporte = reporte.rename(columns={
                    'GRUPO_ALIMENTARIO': 'Grupo', 
                    'PRODUCTO': 'Productos', 
                    'plaza': nombre_pdm, 
                    'externo': 'Tiendas'
                })
                
                reporte['Dif. Precio ($)'] = reporte['Tiendas'] - reporte[nombre_pdm]
                reporte['Dif. Porc. (%)'] = reporte.apply(
                    lambda r: (r['Dif. Precio ($)'] / r[nombre_pdm]) if r[nombre_pdm] != 0 else 0, axis=1
                )
                
                reporte = reporte[['Grupo', 'Productos', nombre_pdm, 'Tiendas', 'Dif. Precio ($)', 'Dif. Porc. (%)']].round(2)

                # Nombre de hoja válido para Excel
                sheet_name = str(plaza)[:31].replace(':', '').replace('/', '')
                reporte.to_excel(writer, sheet_name=sheet_name, index=False)
                
                ws = writer.sheets[sheet_name]
                ws.set_column('A:B', 25)
                ws.set_column('C:E', 15, num_fmt)
                ws.set_column('F:F', 15, percent_fmt)
                
                # Fila de totales con color limitado (B a D)
                idx_total = len(reporte) + 1
                ws.write(idx_total, 1, 'Suma total', total_text_fmt)
                ws.write(idx_total, 2, reporte[nombre_pdm].sum(), total_num_fmt)
                ws.write(idx_total, 3, reporte['Tiendas'].sum(), total_num_fmt)

        print("Archivo generado correctamente: Reporte_Comparativo.xlsx")
    except Exception as e:
        print(f"Error detectado: {e}")

if __name__ == '__main__':
    generar_reportes_completos()
