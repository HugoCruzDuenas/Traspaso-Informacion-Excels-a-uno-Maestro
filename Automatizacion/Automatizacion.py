import os
import glob
import pandas as pd
import numpy as np
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog, messagebox

# --- FUNCIÓN PARA LA INTERFAZ ---
def obtener_rutas():
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal de tkinter
    
    messagebox.showinfo("Paso 1", "Selecciona la CARPETA donde están los archivos Excel de datos (PCTools).")
    ruta_datos = filedialog.askdirectory(title="Seleccionar carpeta de datos")
    
    if not ruta_datos:
        messagebox.showwarning("Cancelado", "No se seleccionó carpeta. Saliendo...")
        return None, None

    messagebox.showinfo("Paso 2", "Selecciona el ARCHIVO Excel Maestro (BD_01122025_macro.xlsm).")
    archivo_maestro = filedialog.askopenfilename(
        title="Seleccionar archivo Maestro",
        filetypes=[("Excel con macros", "*.xlsm"), ("Excel", "*.xlsx")]
    )
    
    if not archivo_maestro:
        messagebox.showwarning("Cancelado", "No se seleccionó el maestro. Saliendo...")
        return None, None
        
    return ruta_datos, archivo_maestro

# --- LÓGICA DE PROCESAMIENTO (Tu código original) ---
mapeo_columnas = {
    'System abbr.': ['SYSTEM', 'ABRÉV', 'SISTEMA'],
    'Supplier': ['SUPPLIER', 'FOURNISSEUR', 'PROVEEDOR'],
    'Software': ['SOFTWARE', 'LOGICIEL', 'SW'],
    'VERSION': ['VERSION', 'VER.'],
    'Measurement/Monitoring Tool': ['MEASUREMENT', 'MONITORING', 'SUPERVISION', 'OUTIL'],
    'Installation / Use Restrictions': ['RESTRICTIONS', 'INSTALLATION', 'LIMITATIONS']
}

def normalizar_extremo(val):
    if pd.isna(val) or str(val).strip().lower() in ['nan', 'none', '', 'null', 'empty']:
        return "EMPTY"
    s = str(val).replace('\n', ' ').replace('\r', ' ').strip()
    if s.endswith('.0'): s = s[:-2]
    return "".join(s.split()).upper()

def obtener_mapa_maestro(hoja):
    mapa = {}
    last_row = hoja.Cells(hoja.Rows.Count, "A").End(-4162).Row
    if last_row < 2: return mapa
    
    for r in range(2, last_row + 1):
        proy = normalizar_extremo(hoja.Cells(r, 1).Value)
        syst = normalizar_extremo(hoja.Cells(r, 3).Value)
        supp = normalizar_extremo(hoja.Cells(r, 2).Value)
        soft = normalizar_extremo(hoja.Cells(r, 4).Value)
        
        if soft != "EMPTY" and proy != "EMPTY":
            clave = (proy, syst, supp, soft)
            mapa[clave] = r
    return mapa

def extraer_con_pandas(ruta):
    try:
        xl = pd.ExcelFile(ruta)
        nombre_hoja = next((s for s in xl.sheet_names if any(x in str(s).upper().replace("_", "").replace(" ", "") 
                           for x in ["PCTOOL", "LCSOFITEM", "SWLIST", "PCTOOLS_BD"])), None)
        if not nombre_hoja: return None
        df = pd.read_excel(ruta, sheet_name=nombre_hoja, header=None)
        
        fila_cab_idx = None
        for i in range(min(len(df), 50)):
            fila_texto = " ".join([str(v).upper() for v in df.iloc[i].fillna('')])
            if "SOFTWARE" in fila_texto:
                fila_cab_idx = i
                break
        
        if fila_cab_idx is None: return None
        
        cabeceras = [str(c).upper().replace('\n', ' ').strip() for c in df.iloc[fila_cab_idx].fillna('')]
        indices = {est: next((idx for idx, c in enumerate(cabeceras) if any(p in c for p in pals)), None) 
                  for est, pals in mapeo_columnas.items()}
        
        temp_df = df.iloc[fila_cab_idx + 1:].copy()
        data_dict = {col: temp_df.iloc[:, idx] if idx is not None else "" for col, idx in indices.items()}
        res = pd.DataFrame(data_dict)
        res = res.replace(r'^\s*$', np.nan, regex=True).replace(['nan', 'None', 'nan.0'], np.nan)
        
        res['System abbr.'] = res['System abbr.'].ffill()
        res['Supplier'] = res['Supplier'].ffill()
        res = res[res['Software'].notna() & (res['Software'].astype(str).str.strip() != '')]
        return res
    except Exception as e:
        print(f"❌ Error leyendo {os.path.basename(ruta)}: {e}")
        return None

# --- EJECUCIÓN ---
if __name__ == "__main__":
    ruta_datos, archivo_maestro = obtener_rutas()
    
    if ruta_datos and archivo_maestro:
        excel_app = None
        try:
            print("🚀 Iniciando motor de Excel...")
            excel_app = win32.gencache.EnsureDispatch('Excel.Application')
            excel_app.Visible = False
            excel_app.DisplayAlerts = False 

            wb_maestro = excel_app.Workbooks.Open(archivo_maestro)
            hoja = wb_maestro.Worksheets("PCTools_BD")

            mapa_existentes = obtener_mapa_maestro(hoja)
            print(f"Registros base cargados: {len(mapa_existentes)}")

            last_row = hoja.Cells(hoja.Rows.Count, "A").End(-4162).Row
            archivos = glob.glob(os.path.join(ruta_datos, "*.xls*"))
            
            total_nuevos = 0
            total_actualizados = 0

            for ruta in archivos:
                nombre_archivo = os.path.basename(ruta)
                if nombre_archivo in archivo_maestro or "~$" in nombre_archivo: continue
                
                id_proyecto_bruto = nombre_archivo[:3].upper()
                id_proyecto_norm = normalizar_extremo(id_proyecto_bruto)
                
                datos = extraer_con_pandas(ruta)
                if datos is not None and not datos.empty:
                    print(f"\n--- Analizando: {nombre_archivo} ---")
                    for _, fila in datos.iterrows():
                        syst_n = normalizar_extremo(fila['System abbr.'])
                        supp_n = normalizar_extremo(fila['Supplier'])
                        soft_n = normalizar_extremo(fila['Software'])
                        ver_n = str(fila['VERSION']).strip() if pd.notna(fila['VERSION']) else ""
                        
                        clave = (id_proyecto_norm, syst_n, supp_n, soft_n)

                        if clave in mapa_existentes:
                            fila_idx = mapa_existentes[clave]
                            ver_actual = str(hoja.Cells(fila_idx, 5).Value).strip()
                            if normalizar_extremo(ver_n) != normalizar_extremo(ver_actual):
                                print(f"  [ACTUALIZADO] {fila['Software']}: {ver_actual} -> {ver_n}")
                                hoja.Cells(fila_idx, 5).Value = ver_n
                                total_actualizados += 1
                        else:
                            last_row += 1
                            hoja.Cells(last_row, 1).Value = id_proyecto_bruto
                            hoja.Cells(last_row, 2).Value = str(fila['Supplier']) if pd.notna(fila['Supplier']) else ""
                            hoja.Cells(last_row, 3).Value = str(fila['System abbr.']) if pd.notna(fila['System abbr.']) else ""
                            hoja.Cells(last_row, 4).Value = str(fila['Software'])
                            hoja.Cells(last_row, 5).Value = ver_n
                            hoja.Cells(last_row, 6).Value = str(fila['Measurement/Monitoring Tool']) if pd.notna(fila['Measurement/Monitoring Tool']) else ""
                            hoja.Cells(last_row, 7).Value = str(fila['Installation / Use Restrictions']) if pd.notna(fila['Installation / Use Restrictions']) else ""
                            mapa_existentes[clave] = last_row
                            total_nuevos += 1

            wb_maestro.Save()
            messagebox.showinfo("Fin", f"Proceso terminado.\n\nActualizados: {total_actualizados}\nNuevos: {total_nuevos}")

        except Exception as e:
            messagebox.showerror("Error crítico", f"Ocurrió un error: {e}")
        finally:
            if excel_app:
                try:
                    wb_maestro.Close()
                    excel_app.Quit()
                except: pass