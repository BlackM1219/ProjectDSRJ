
# -*- coding: utf-8 -*-
"""
Formulario DISAGRO - Sistema completo con integración a Teams
Versión con Power Automate Webhook + Conversión PDF desde Python
 
INSTRUCCIONES:
1. Instala dependencias: pip install openpyxl pywin32 requests
2. Configura tu WEBHOOK_URL (línea ~30)
3. Asegúrate de tener el archivo plantilla.xlsx
4. ¡Ejecuta y usa!
 
FLUJO:
  1. Buscar usuario en Outlook
  2. Rellenar datos del formulario
  3. [✅ Generar Excel]         → guarda .xlsx localmente
  4. [📤 Convertir PDF y Enviar] → convierte a .pdf y envía por Power Automate
"""
 
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
import win32com.client
from typing import Optional, Dict
from dataclasses import dataclass, field
import re
from datetime import datetime
import os
import requests
import base64
 
TEMPLATE_PATH = "plantilla.xlsx"
 
# ============================================
# 🔧 CONFIGURA AQUÍ TU URL DEL WEBHOOK
# ============================================
WEBHOOK_URL = "https://default93a2cd4474774ba69ca3e4e7aaf9c2.dd.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/0f4c9b26740b401c9681e227af854f4e/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=wfREfiqzfGGgVHBoQajhWsotqJs7m6_FbPD8zq1mnwc"
# ============================================
 
 
# ============================================================
# 🔄 CONVERSIÓN EXCEL → PDF
# ============================================================
def convertir_excel_a_pdf(excel_path: str) -> Optional[str]:
    """
    Convierte un archivo .xlsx a .pdf usando Excel instalado en Windows.
    Retorna la ruta del PDF generado, o None si falla.
    """
    excel_app = None
    wb = None
    try:
        excel_abs = os.path.abspath(excel_path)
        pdf_path  = excel_abs.replace('.xlsx', '.pdf')
 
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible       = False
        excel_app.DisplayAlerts = False
 
        wb = excel_app.Workbooks.Open(excel_abs)
        wb.ExportAsFixedFormat(0, pdf_path)   # 0 = xlTypePDF
        return pdf_path
 
    except Exception as e:
        messagebox.showerror(
            "❌ Error al convertir a PDF",
            f"No se pudo convertir el Excel a PDF.\n\nDetalle: {str(e)}"
        )
        return None
    finally:
        try:
            if wb:        wb.Close(False)
        except: pass
        try:
            if excel_app: excel_app.Quit()
        except: pass
 
 
# ============================================================
# 📦 ESTRUCTURA DE DATOS
# ============================================================
@dataclass
class UserData:
    alias: str = ""
    full_name: str = ""
    first_name: str = ""
    second_name: str = ""
    first_surname: str = ""
    second_surname: str = ""
    correo: str = ""
    puesto: str = ""
    empresa: str = ""
    depart: str = ""
    ticket: str = ""
    fecha: str = ""
    localidad: str = "GT"
    pais: str = "GT"
    pc_type: str = ""
    marca: str = ""
    model: str = ""
    st: str = ""
    fecha_compra: str = ""
    disco_duro: str = ""
    memoria_ram: str = ""
    ip: str = ""
    hostname: str = ""
    estado: str = "OK"
    more_equipment: list = field(default_factory=list)
    temp_equipos: list = field(default_factory=list)
    sistema_operativo: str = ""
    version_so: str = ""
    ofimática: str = ""
    version_office: str = ""
    reported_problem: str = ""
    diagnosis: str = ""
    documentation: str = ""
    tecnico: str = "Angel Lopez"
    usuario_firma: str = ""
    jefe_responsable: str = "Eder Morales"
    observaciones: str = ""
    form_types: list = field(default_factory=list)
 
 
# ============================================================
# 📤 TEAMS WEBHOOK
# ============================================================
class TeamsWebhook:
    def __init__(self, webhook_url: str):
        self.webhook_url = webhook_url
 
    def enviar_archivo(self, file_path: str, email_destinatario: str, mensaje: str = "") -> bool:
        webhook_url = (self.webhook_url or "").strip()
        if not webhook_url or "PEGA_TU_URL_AQUI" in webhook_url or not webhook_url.startswith("http"):
            messagebox.showerror("⚠️ Webhook no configurado",
                "Configura el WEBHOOK_URL en el código.")
            return False
        try:
            with open(file_path, 'rb') as f:
                file_base64 = base64.b64encode(f.read()).decode('utf-8')
            payload = {
                "email":             email_destinatario,
                "mensaje":           mensaje or f"📋 Nuevo formulario: {os.path.basename(file_path)}",
                "archivo_nombre":    os.path.basename(file_path),
                "archivo_contenido": file_base64
            }
            response = requests.post(webhook_url, json=payload,
                headers={"Content-Type": "application/json"}, timeout=30)
            if response.status_code in [200, 202, 204]:
                messagebox.showinfo("✅ Enviado exitosamente",
                    f"Archivo enviado a {email_destinatario}\n\n"
                    f"El destinatario recibirá el mensaje en Teams en unos segundos.")
                return True
            else:
                messagebox.showerror("❌ Error al enviar",
                    f"Código {response.status_code}\n{response.text[:200]}")
                return False
        except FileNotFoundError:
            messagebox.showerror("❌ Error", f"No se encontró el archivo:\n{file_path}")
        except requests.exceptions.Timeout:
            messagebox.showerror("⏱️ Tiempo agotado", "El servidor tardó demasiado. Intenta de nuevo.")
        except requests.exceptions.ConnectionError:
            messagebox.showerror("🌐 Error de conexión", "Verifica tu conexión a internet.")
        except Exception as e:
            messagebox.showerror("❌ Error", f"Error al enviar:\n{str(e)}")
        return False
 
 
# ============================================================
# 🔍 BÚSQUEDA EN OUTLOOK
# ============================================================
class OutlookSearcher:
    @staticmethod
    def buscar_usuario(alias: str) -> Optional[Dict[str, str]]:
        alias = alias.lower().split("@")[0].strip()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            ns  = outlook.GetNamespace("MAPI")
            gal = ns.AddressLists["Lista global de direcciones"]
            for i in range(gal.AddressEntries.Count):
                entry = gal.AddressEntries.Item(i + 1)
                try:
                    ex     = entry.GetExchangeUser()
                    correo = ex.PrimarySmtpAddress.lower()
                    if correo.startswith(alias + "@"):
                        parts          = ex.Name.split()
                        first_name     = parts[0] if parts else ""
                        second_name    = ""
                        first_surname  = ""
                        second_surname = ""
                        if len(parts) >= 3:
                            first_surname  = parts[-2]
                            second_surname = parts[-1]
                            if len(parts) > 3:
                                second_name = " ".join(parts[1:-2])
                        elif len(parts) == 2:
                            first_surname = parts[1]
                        return {
                            "alias": alias, "full_name": ex.Name,
                            "first_name": first_name, "second_name": second_name,
                            "first_surname": first_surname, "second_surname": second_surname,
                            "correo": correo, "puesto": ex.JobTitle or "",
                            "empresa": ex.CompanyName or "", "depart": ex.Department or ""
                        }
                except AttributeError:
                    continue
            return None
        except Exception as e:
            messagebox.showerror("Error COM", f"Error al conectar con Outlook:\n{str(e)}")
            return None
 
 
# ============================================================
# 📊 GENERADOR DE EXCEL
# ============================================================
class ExcelGenerator:
    FORM_TYPES = [
        "FECHA DE DEVOLUCIÓN DE EQUIPO",
        "HOJA DE RESPONSABILIDAD",
        "SOPORTE TÉCNICO",
        "PASE DE SALIDA",
        "DICTAMEN TÉCNICO",
        "RECEPCIÓN DE EQUIPO",
        "PRÉSTAMO DE EQUIPO"
    ]
 
    def __init__(self, template_path: str = TEMPLATE_PATH):
        self.template_path = template_path
 
    def generar(self, data: UserData, output_dir: str = ".") -> Optional[str]:
        if not data.ticket or not data.alias:
            messagebox.showwarning("Faltan datos",
                "Debes buscar un usuario y completar al menos el número de ticket.")
            return None
        try:
            wb = load_workbook(self.template_path)
            ws = wb.active
            if not data.fecha:
                data.fecha = datetime.now().strftime("%d/%m/%Y")
            self._reemplazar_etiquetas(ws, self._preparar_datos(data))
            self._marcar_checkboxes(ws, data)
            equipos = data.temp_equipos if data.temp_equipos else data.more_equipment
            if equipos:
                self._llenar_tabla_equipos(ws, equipos)
            safe_alias  = re.sub(r'[^\w\-]', '_', data.alias)
            safe_ticket = re.sub(r'[^\w\-]', '_', data.ticket)
            output_file = os.path.join(output_dir, f"{safe_ticket}_{safe_alias}.xlsx")
            wb.save(output_file)
            return output_file
        except FileNotFoundError:
            messagebox.showerror("Error", f"No se encontró la plantilla: {self.template_path}")
        except Exception as e:
            messagebox.showerror("Error al generar", f"Error al crear la plantilla:\n{str(e)}")
        return None
 
    def _preparar_datos(self, data: UserData) -> Dict[str, str]:
        return {
            "ticket": data.ticket, "fecha": data.fecha,
            "alias": data.alias, "first_name": data.first_name,
            "second_name": data.second_name, "first_surname": data.first_surname,
            "second_surname": data.second_surname, "puesto": data.puesto,
            "depart": data.depart, "empresa": data.empresa,
            "localidad": data.localidad, "pais": data.pais, "correo": data.correo,
            "pc_type": data.pc_type, "marca": data.marca, "model": data.model,
            "st": data.st, "fecha_compra": data.fecha_compra,
            "disco_duro": data.disco_duro, "memoria_ram": data.memoria_ram,
            "ip": data.ip, "hostname": data.hostname, "estado": data.estado,
            "sistema_operativo": data.sistema_operativo, "version_so": data.version_so,
            "ofimatica": data.ofimática, "version_office": data.version_office,
            "reported_problem": data.reported_problem, "diagnosis": data.diagnosis,
            "documentation": data.documentation, "tecnico": data.tecnico,
            "usuario_firma": data.usuario_firma or data.full_name or data.alias,
            "jefe_responsable": data.jefe_responsable, "observaciones": data.observaciones,
        }
 
    def _reemplazar_etiquetas(self, ws, replacements: Dict[str, str]):
        pattern = re.compile(r'\{\{(\w+)\}\}')
        for row_idx in range(1, ws.max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                try:
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if isinstance(cell, MergedCell): continue
                    if cell.value and isinstance(cell.value, str) and "{{" in cell.value:
                        new_value = cell.value
                        for tag in pattern.findall(cell.value):
                            if tag in replacements:
                                new_value = new_value.replace(f"{{{{{tag}}}}}", str(replacements[tag]))
                        if new_value != cell.value:
                            self._escribir_celda_segura(ws, row_idx, col_idx, new_value)
                except Exception:
                    continue
 
    def _escribir_celda_segura(self, ws, row: int, col: int, value):
        try:
            celda = ws.cell(row=row, column=col)
            if isinstance(celda, MergedCell):
                for mr in ws.merged_cells.ranges:
                    if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
                        ws.cell(row=mr.min_row, column=mr.min_col, value=value)
                        return
            else:
                celda.value = value
        except Exception:
            try: ws.cell(row=row, column=col).value = value
            except: pass
 
    def _col_letter_to_index(self, letter: str) -> int:
        idx = 0
        for ch in letter.upper():
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        return idx
 
    def _marcar_checkboxes(self, ws, data: UserData):
        form_checkboxes = {
            "HOJA DE RESPONSABILIDAD": ("B", 6), "SOPORTE TÉCNICO": ("M", 6),
            "SOPORTE TECNICO": ("M", 6), "PASE DE SALIDA": ("V", 6),
            "DICTAMEN TÉCNICO": ("B", 7), "DICTAMEN TECNICO": ("B", 7),
            "RECEPCIÓN DE EQUIPO": ("M", 7), "RECEPCION DE EQUIPO": ("M", 7),
            "PRÉSTAMO DE EQUIPO": ("V", 7), "PRESTAMO DE EQUIPO": ("V", 7),
        }
        for ft in data.form_types:
            if not ft: continue
            key = ft.strip().upper().replace("Á","A").replace("É","E").replace("Í","I").replace("Ó","O").replace("Ú","U")
            if key in form_checkboxes:
                col_l, row = form_checkboxes[key]
                try: self._escribir_celda_segura(ws, row, self._col_letter_to_index(col_l), "X")
                except Exception as e: print(f"Error checkbox {ft}: {e}")
        if data.sistema_operativo:
            so = data.sistema_operativo.upper()
            if "WINDOWS" in so: self._escribir_celda_segura(ws, 38, self._col_letter_to_index("J"), "X")
            if "MAC"     in so: self._escribir_celda_segura(ws, 38, self._col_letter_to_index("N"), "X")
        if data.ofimática and ("OFFICE" in data.ofimática.upper() or "MS" in data.ofimática.upper()):
            self._escribir_celda_segura(ws, 39, self._col_letter_to_index("J"), "X")
 
    def _llenar_tabla_equipos(self, ws, equipos: list):
        fila_inicio = 22
        cols = {
            "equipo": self._col_letter_to_index("C"),
            "marca":  self._col_letter_to_index("M"),
            "modelo": self._col_letter_to_index("U"),
            "serie":  self._col_letter_to_index("AC"),
            "estado": self._col_letter_to_index("AO"),
        }
        for i, eq in enumerate(equipos):
            r = fila_inicio + i
            try:
                for key, col in cols.items():
                    self._escribir_celda_segura(ws, r, col, eq.get(key, "" if key != "estado" else "OK"))
            except Exception as e:
                print(f"Error llenando equipo {i}: {e}")
 
 
# ============================================================
# 🖥️ APLICACIÓN PRINCIPAL
# ============================================================
class FormularioApp:
    def __init__(self, root):
        self.root          = root
        self.root.title("Formulario DISAGRO - Departamento de Sistemas")
        self.root.geometry("900x700")
        self.data          = UserData()
        self.excel_gen     = ExcelGenerator()
        self.teams_webhook = TeamsWebhook(WEBHOOK_URL)
        self.widgets       = {}
        self.ultimo_excel  = None   # ruta del .xlsx generado
        self._crear_interfaz()
 
    def _crear_interfaz(self):
        # Título
        title_frame = tk.Frame(self.root, bg="#2c3e50", padx=10, pady=10)
        title_frame.grid(row=0, column=0, sticky="ew")
        tk.Label(title_frame,
            text="SISTEMA DE FORMULARIOS - DEPARTAMENTO DE SISTEMAS",
            font=("Arial", 14, "bold"), bg="#2c3e50", fg="white").pack()
 
        # Búsqueda
        search_frame = tk.LabelFrame(self.root, text="1. Buscar Usuario", padx=10, pady=10)
        search_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        tk.Label(search_frame, text="Alias o correo:").grid(row=0, column=0, sticky="w")
        self.widgets['alias_entry'] = tk.Entry(search_frame, width=40)
        self.widgets['alias_entry'].grid(row=0, column=1, padx=5)
        self.widgets['alias_entry'].bind('<Return>', lambda e: self._buscar_usuario())
        tk.Button(search_frame, text="🔍 Buscar en Outlook",
            command=self._buscar_usuario, bg="#3498db", fg="white", padx=10).grid(row=0, column=2)
 
        self._crear_treeview()
 
        # Botones
        button_frame = tk.LabelFrame(self.root, text="Acciones", padx=10, pady=10)
        button_frame.grid(row=3, column=0, padx=10, pady=5)
 
        # Fila 0: Rellenar datos
        self.widgets['btn_rellenar'] = tk.Button(
            button_frame, text="📝 Rellenar Datos del Formulario",
            command=self._abrir_formulario_datos, state="disabled",
            bg="#e67e22", fg="white", padx=20, pady=10, font=("Arial", 10, "bold"))
        self.widgets['btn_rellenar'].grid(row=0, column=0, columnspan=3, pady=(0, 10))
 
        # Fila 1: Dos botones del flujo
        self.widgets['btn_excel'] = tk.Button(
            button_frame, text="✅ 1. Generar Excel",
            command=self._generar_excel, state="disabled",
            bg="#27ae60", fg="white", padx=20, pady=10, font=("Arial", 10, "bold"))
        self.widgets['btn_excel'].grid(row=1, column=0, padx=5)
 
        tk.Label(button_frame, text="→", font=("Arial", 16)).grid(row=1, column=1, padx=10)
 
        self.widgets['btn_enviar'] = tk.Button(
            button_frame, text="📤 2. Convertir PDF y Enviar a Teams",
            command=self._convertir_y_enviar, state="disabled",
            bg="#0078d4", fg="white", padx=20, pady=10, font=("Arial", 10, "bold"))
        self.widgets['btn_enviar'].grid(row=1, column=2, padx=5)
 
        # Fila 2: Estado
        self.widgets['lbl_estado'] = tk.Label(
            button_frame, text="⬜ Sin archivo generado",
            font=("Arial", 9), fg="gray")
        self.widgets['lbl_estado'].grid(row=2, column=0, columnspan=3, pady=(10, 0))

        tk.Button(
            button_frame, text="Limpiar todos los datos",
            command=self._limpiar_datos, bg="#c0392b", fg="white",
            padx=12, pady=6).grid(row=3, column=0, columnspan=3, pady=(12, 0))
 
    def _crear_treeview(self):
        tree_frame = tk.LabelFrame(self.root, text="2. Usuario Encontrado", padx=10, pady=5)
        tree_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=5)
        columns = ("Nombre", "Correo", "Puesto", "Empresa", "Departamento")
        self.widgets['tree'] = ttk.Treeview(tree_frame, columns=columns, show="headings", height=3)
        for col in columns:
            self.widgets['tree'].heading(col, text=col)
            self.widgets['tree'].column(col, width=160)
        self.widgets['tree'].pack(fill="both", expand=True)
 
    def _actualizar_estado(self, texto: str, color: str = "gray"):
        self.widgets['lbl_estado'].config(text=texto, fg=color)

    def _limpiar_datos(self):
        if not messagebox.askyesno(
                "Limpiar datos", "Se borrarán los datos cargados en la aplicación. ¿Continuar?"):
            return

        for window in self.root.winfo_children():
            if isinstance(window, tk.Toplevel):
                window.destroy()

        self.data = UserData()
        self.ultimo_excel = None
        self.widgets['alias_entry'].delete(0, tk.END)
        self.widgets['tree'].delete(*self.widgets['tree'].get_children())
        self.widgets['btn_rellenar'].config(state="disabled")
        self.widgets['btn_excel'].config(state="disabled")
        self.widgets['btn_enviar'].config(state="disabled")
        self._actualizar_estado("Sin archivo generado")
 
    def _buscar_usuario(self):
        alias = self.widgets['alias_entry'].get().strip()
        if not alias:
            messagebox.showwarning("Campo vacío", "Ingresa un alias o correo.")
            return
        self.widgets['tree'].delete(*self.widgets['tree'].get_children())
        resultado = OutlookSearcher.buscar_usuario(alias)
        if resultado:
            for key, value in resultado.items():
                setattr(self.data, key, value)
            self.widgets['tree'].insert("", "end", values=(
                resultado["full_name"], resultado["correo"],
                resultado["puesto"], resultado["empresa"], resultado["depart"]))
            self.widgets['btn_rellenar'].config(state="normal")
            self.widgets['btn_excel'].config(state="normal")
            messagebox.showinfo("✓ Usuario encontrado", f"Usuario {resultado['full_name']} cargado.")
        else:
            self.widgets['btn_rellenar'].config(state="disabled")
            self.widgets['btn_excel'].config(state="disabled")
            messagebox.showwarning("No encontrado", f"No se encontró el usuario: {alias}")
 
    def _abrir_formulario_datos(self):
        if not self.data.alias:
            messagebox.showwarning("Busca primero", "Primero debes buscar un usuario.")
            return
 
        win = tk.Toplevel(self.root)
        win.title(f"Datos del Formulario - {self.data.full_name}")
        win.geometry("800x900")
 
        canvas           = tk.Canvas(win)
        scrollbar        = tk.Scrollbar(win, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        scrollable_frame.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
 
        entries = {}
        row     = 0
 
        tk.Label(scrollable_frame, text="DATOS DEL FORMULARIO", font=("Arial", 12, "bold"),
                 bg="#2c3e50", fg="white").grid(row=row, column=0, columnspan=2, sticky="ew", pady=(0,10))
        row += 1
 
        for lbl, key, default in [
            ("No. de Ticket *", "ticket", self.data.ticket),
            ("Fecha",           "fecha",  datetime.now().strftime("%d/%m/%Y")),
        ]:
            tk.Label(scrollable_frame, text=lbl).grid(row=row, column=0, sticky="e", padx=5, pady=5)
            entries[key] = tk.Entry(scrollable_frame, width=50)
            entries[key].grid(row=row, column=1, sticky="w", padx=5, pady=5)
            entries[key].insert(0, default)
            row += 1
 
        tk.Label(scrollable_frame, text="TIPO DE FORMULARIO", font=("Arial", 11, "bold"),
                 bg="#34495e", fg="white").grid(row=row, column=0, columnspan=2, sticky="ew", pady=(15,5))
        row += 1
        tk.Label(scrollable_frame, text="Selecciona tipo(s) *").grid(row=row, column=0, sticky="ne", padx=5, pady=5)
        form_frame = tk.Frame(scrollable_frame)
        form_frame.grid(row=row, column=1, sticky="w", padx=5, pady=5)
        form_vars = {}
        for i, ft in enumerate(ExcelGenerator.FORM_TYPES):
            var = tk.BooleanVar(value=ft in self.data.form_types)
            tk.Checkbutton(form_frame, text=ft, variable=var).grid(row=i, column=0, sticky="w")
            form_vars[ft] = var
        entries['form_types'] = form_vars
        row += 1
 
        tk.Label(scrollable_frame, text="A. DATOS DEL COLABORADOR", font=("Arial", 11, "bold"),
                 bg="#34495e", fg="white").grid(row=row, column=0, columnspan=2, sticky="ew", pady=(15,5))
        row += 1
        tk.Label(scrollable_frame, text="(Ya cargados desde Outlook)", font=("Arial", 9, "italic"),
                 fg="gray").grid(row=row, column=0, columnspan=2, pady=2)
        row += 1
        for lbl, key in [("Localidad", "localidad"), ("País", "pais")]:
            tk.Label(scrollable_frame, text=lbl).grid(row=row, column=0, sticky="e", padx=5, pady=5)
            entries[key] = tk.Entry(scrollable_frame, width=50)
            entries[key].grid(row=row, column=1, sticky="w", padx=5, pady=5)
            entries[key].insert(0, getattr(self.data, key, ""))
            row += 1
 
        tk.Label(scrollable_frame, text="B. HARDWARE", font=("Arial", 11, "bold"),
                 bg="#34495e", fg="white").grid(row=row, column=0, columnspan=2, sticky="ew", pady=(15,5))
        row += 1
        campos_hardware = [
            ("Tipo de Computadora *", "pc_type"), ("Marca *", "marca"),
            ("Modelo", "model"), ("Service Tag (ST) *", "st"),
            ("Fecha de Compra", "fecha_compra"), ("Disco Duro", "disco_duro"),
            ("Memoria RAM", "memoria_ram"), ("IP", "ip"), ("Host Name", "hostname"),
        ]
        for lbl, key in campos_hardware:
            tk.Label(scrollable_frame, text=lbl).grid(row=row, column=0, sticky="e", padx=5, pady=5)
            entries[key] = tk.Entry(scrollable_frame, width=50)
            entries[key].grid(row=row, column=1, sticky="w", padx=5, pady=5)
            entries[key].insert(0, getattr(self.data, key, ""))
            row += 1
 
        tk.Label(scrollable_frame, text="Estado del equipo").grid(row=row, column=0, sticky="e", padx=5, pady=5)
        entries['estado'] = ttk.Combobox(scrollable_frame, values=["OK","MALO","REGULAR"],
                                          state="readonly", width=47)
        entries['estado'].grid(row=row, column=1, sticky="w", padx=5, pady=5)
        entries['estado'].set(self.data.estado)
        row += 1
 
        tk.Label(scrollable_frame, text="C. EQUIPOS ADICIONALES", font=("Arial", 11, "bold"),
                 bg="#34495e", fg="white").grid(row=row, column=0, columnspan=2, sticky="ew", pady=(15,5))
        row += 1
        eq_frame = tk.Frame(scrollable_frame, relief="sunken", bd=1)
        eq_frame.grid(row=row, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        equipos_list = tk.Listbox(eq_frame, height=5, width=80)
        equipos_list.pack(side="left", fill="both", expand=True)
        eq_sb = tk.Scrollbar(eq_frame, command=equipos_list.yview)
        eq_sb.pack(side="right", fill="y")
        equipos_list.config(yscrollcommand=eq_sb.set)
        for eq in self.data.more_equipment:
            equipos_list.insert("end", f"{eq.get('equipo','')} - {eq.get('marca','')} {eq.get('modelo','')} - {eq.get('serie','')}")
 
        btn_eq = tk.Frame(scrollable_frame)
        btn_eq.grid(row=row+1, column=0, columnspan=2)
 
        def agregar_equipo():
            eq_win = tk.Toplevel(win)
            eq_win.title("Agregar Equipo Adicional")
            eq_win.geometry("450x300")
            eq_entries = {}
            for i, (lbl, key) in enumerate(zip(
                ["Equipo:","Marca:","Modelo:","Serie:","Estado:"],
                ["equipo","marca","modelo","serie","estado"]
            )):
                tk.Label(eq_win, text=lbl).grid(row=i, column=0, sticky="e", padx=5, pady=5)
                if key == "estado":
                    e = ttk.Combobox(eq_win, values=["OK","MALO","REGULAR"], state="readonly", width=28)
                    e.set("OK")
                else:
                    e = tk.Entry(eq_win, width=30)
                e.grid(row=i, column=1, padx=5, pady=5)
                eq_entries[key] = e
 
            def guardar_equipo():
                eq_data = {k: v.get().strip() for k, v in eq_entries.items()}
                if eq_data["equipo"]:
                    equipos_list.insert("end", f"{eq_data['equipo']} - {eq_data['marca']} {eq_data['modelo']} - {eq_data['serie']}")
                    self.data.temp_equipos.append(eq_data)
                    eq_win.destroy()
                else:
                    messagebox.showwarning("Campo vacío", "El nombre del equipo es obligatorio")
 
            tk.Button(eq_win, text="💾 Guardar", command=guardar_equipo,
                      bg="#27ae60", fg="white").grid(row=5, column=0, columnspan=2, pady=10)
 
        def eliminar_equipo():
            sel = equipos_list.curselection()
            if not sel: return
            idx = sel[0]
            equipos_list.delete(idx)
            if idx < len(self.data.temp_equipos):
                self.data.temp_equipos.pop(idx)
 
        tk.Button(btn_eq, text="➕ Agregar",  command=agregar_equipo, bg="#3498db", fg="white").grid(row=0, column=0, padx=5, pady=5)
        tk.Button(btn_eq, text="➖ Eliminar", command=eliminar_equipo, bg="#c0392b", fg="white").grid(row=0, column=1, padx=5, pady=5)
        row += 3
 
        tk.Label(scrollable_frame, text="D. SOFTWARE", font=("Arial", 11, "bold"),
                 bg="#34495e", fg="white").grid(row=row, column=0, columnspan=2, sticky="ew", pady=(15,5))
        row += 1
 
        tk.Label(scrollable_frame, text="Sistema operativo").grid(row=row, column=0, sticky="ne", padx=5, pady=5)
        so_frame = tk.Frame(scrollable_frame)
        so_frame.grid(row=row, column=1, sticky="w", padx=5, pady=5)
        so_vars = {}
        for i, opt in enumerate(["Windows", "Mac OS"]):
            var = tk.BooleanVar(value=opt in self.data.sistema_operativo)
            tk.Checkbutton(so_frame, text=opt, variable=var).grid(row=0, column=i, sticky="w", padx=10)
            so_vars[opt] = var
        entries['sistema_operativo_vars'] = so_vars
        row += 1
 
        tk.Label(scrollable_frame, text="Versión SO").grid(row=row, column=0, sticky="e", padx=5, pady=5)
        entries['version_so'] = tk.Entry(scrollable_frame, width=50)
        entries['version_so'].grid(row=row, column=1, sticky="w", padx=5, pady=5)
        entries['version_so'].insert(0, self.data.version_so)
        row += 1
 
        tk.Label(scrollable_frame, text="Ofimática").grid(row=row, column=0, sticky="ne", padx=5, pady=5)
        ofi_frame = tk.Frame(scrollable_frame)
        ofi_frame.grid(row=row, column=1, sticky="w", padx=5, pady=5)
        ofi_vars = {}
        for i, opt in enumerate(["MS Office", "Open Office"]):
            var = tk.BooleanVar(value=opt in self.data.ofimática)
            tk.Checkbutton(ofi_frame, text=opt, variable=var).grid(row=0, column=i, sticky="w", padx=10)
            ofi_vars[opt] = var
        entries['ofimatica_vars'] = ofi_vars
        row += 1
 
        tk.Label(scrollable_frame, text="Versión Office").grid(row=row, column=0, sticky="e", padx=5, pady=5)
        entries['version_office'] = tk.Entry(scrollable_frame, width=50)
        entries['version_office'].grid(row=row, column=1, sticky="w", padx=5, pady=5)
        entries['version_office'].insert(0, self.data.version_office)
        row += 1
 
        tk.Label(scrollable_frame, text="E. PROBLEMA REPORTADO", font=("Arial", 11, "bold"),
                 bg="#34495e", fg="white").grid(row=row, column=0, columnspan=2, sticky="ew", pady=(15,5))
        row += 1
        tk.Label(scrollable_frame, text="Problema reportado:").grid(row=row, column=0, sticky="ne", padx=5, pady=5)
        entries['reported_problem'] = tk.Text(scrollable_frame, width=50, height=4)
        entries['reported_problem'].grid(row=row, column=1, sticky="w", padx=5, pady=5)
        if self.data.reported_problem:
            entries['reported_problem'].insert("1.0", self.data.reported_problem)
        row += 1
 
        tk.Label(scrollable_frame, text="F. DIAGNÓSTICO", font=("Arial", 11, "bold"),
                 bg="#34495e", fg="white").grid(row=row, column=0, columnspan=2, sticky="ew", pady=(15,5))
        row += 1
        tk.Label(scrollable_frame, text="Diagnóstico:").grid(row=row, column=0, sticky="ne", padx=5, pady=5)
        entries['diagnosis'] = tk.Text(scrollable_frame, width=50, height=6)
        entries['diagnosis'].grid(row=row, column=1, sticky="w", padx=5, pady=5)
        if self.data.diagnosis:
            entries['diagnosis'].insert("1.0", self.data.diagnosis)
        row += 1
 
        tk.Label(scrollable_frame, text="Documentación/Repuestos:").grid(row=row, column=0, sticky="ne", padx=5, pady=5)
        entries['documentation'] = tk.Text(scrollable_frame, width=50, height=3)
        entries['documentation'].grid(row=row, column=1, sticky="w", padx=5, pady=5)
        if self.data.documentation:
            entries['documentation'].insert("1.0", self.data.documentation)
        row += 1
 
        def guardar_formulario():
            if not entries['ticket'].get().strip():
                messagebox.showwarning("Faltan datos", "El número de ticket es obligatorio")
                return
            self.data.ticket     = entries['ticket'].get().strip()
            self.data.fecha      = entries['fecha'].get().strip()
            self.data.form_types = [k for k, v in entries['form_types'].items() if v.get()]
            self.data.localidad  = entries['localidad'].get().strip()
            self.data.pais       = entries['pais'].get().strip()
            for _, key in campos_hardware:
                setattr(self.data, key, entries[key].get().strip())
            self.data.estado = entries['estado'].get().strip()
            self.data.sistema_operativo = ", ".join(so for so, v in entries['sistema_operativo_vars'].items() if v.get())
            self.data.version_so        = entries['version_so'].get().strip()
            self.data.ofimática         = ", ".join(ofi for ofi, v in entries['ofimatica_vars'].items() if v.get())
            self.data.version_office    = entries['version_office'].get().strip()
            self.data.reported_problem  = entries['reported_problem'].get("1.0", "end").strip()
            self.data.diagnosis         = entries['diagnosis'].get("1.0", "end").strip()
            self.data.documentation     = entries['documentation'].get("1.0", "end").strip()
            messagebox.showinfo("Guardado", "Datos del formulario guardados en memoria.")
            win.destroy()
 
        tk.Button(scrollable_frame, text="💾 Guardar formulario", command=guardar_formulario,
                  bg="#27ae60", fg="white", padx=10, pady=8,
                  font=("Arial", 10, "bold")).grid(row=row, column=0, columnspan=2, pady=15)
 
    # ----------------------------------------------------------
    # PASO 1: GENERAR EXCEL
    # ----------------------------------------------------------
    def _generar_excel(self):
        if not self.data.alias:
            messagebox.showwarning("Busca primero", "Primero debes buscar un usuario.")
            return
        if not self.data.ticket:
            messagebox.showwarning("Falta ticket", "Completa el número de ticket en 'Rellenar Datos'.")
            return
 
        out_dir = filedialog.askdirectory(title="Selecciona carpeta para guardar", initialdir=os.getcwd())
        if not out_dir:
            if not messagebox.askyesno("Continuar", "No seleccionaste carpeta. ¿Guardar en la carpeta actual?"):
                return
            out_dir = os.getcwd()
 
        excel_file = self.excel_gen.generar(self.data, output_dir=out_dir)
        if excel_file:
            self.ultimo_excel = excel_file
            self.widgets['btn_enviar'].config(state="normal")
            self._actualizar_estado(f"✅ Excel listo: {os.path.basename(excel_file)}", "#27ae60")
            self.data.temp_equipos = []
            messagebox.showinfo("✅ Excel guardado", f"Archivo guardado:\n{excel_file}")
 
    # ----------------------------------------------------------
    # PASO 2: CONVERTIR A PDF Y ENVIAR
    # ----------------------------------------------------------
    def _convertir_y_enviar(self):
        if not self.ultimo_excel or not os.path.exists(self.ultimo_excel):
            messagebox.showwarning("Sin Excel", "Primero genera el Excel (Paso 1).")
            return
 
        # Ventana para pedir email y mensaje
        teams_win = tk.Toplevel(self.root)
        teams_win.title("Convertir PDF y Enviar a Teams")
        teams_win.geometry("500x380")
        teams_win.resizable(False, False)
 
        main_frame = tk.Frame(teams_win, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)
 
        tk.Label(main_frame, text="📤 Convertir PDF y Enviar a Teams",
                 font=("Arial", 13, "bold"), fg="#0078d4").pack(pady=(0, 15))
 
        tk.Label(main_frame, text="Excel a convertir:", font=("Arial", 10, "bold")).pack(anchor="w")
        tk.Label(main_frame, text=os.path.basename(self.ultimo_excel),
                 font=("Arial", 9), fg="gray").pack(anchor="w", pady=(0, 15))
 
        tk.Label(main_frame, text="Email del destinatario:", font=("Arial", 10, "bold")).pack(anchor="w")
        email_entry = tk.Entry(main_frame, width=50, font=("Arial", 10))
        email_entry.pack(fill="x", pady=(5, 15))
        if self.data.correo:
            email_entry.insert(0, self.data.correo)
 
        tk.Label(main_frame, text="Mensaje (opcional):", font=("Arial", 10, "bold")).pack(anchor="w")
        mensaje_text = tk.Text(main_frame, width=50, height=3, font=("Arial", 9))
        mensaje_text.pack(fill="x", pady=(5, 15))
        mensaje_text.insert("1.0", f"📋 Formulario de {self.data.full_name} - Ticket #{self.data.ticket}")
 
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill="x", pady=(10, 0))
 
        def enviar():
            email = email_entry.get().strip()
            if not email or "@" not in email:
                messagebox.showwarning("Email inválido", "Ingresa un email válido.")
                return
 
            mensaje = mensaje_text.get("1.0", "end").strip()
            teams_win.config(cursor="wait")
            teams_win.update()
 
            # 1️⃣ Convertir a PDF
            self._actualizar_estado("⏳ Convirtiendo a PDF...", "#e67e22")
            self.root.update()
            pdf_file = convertir_excel_a_pdf(self.ultimo_excel)
 
            if not pdf_file or not os.path.exists(pdf_file):
                teams_win.config(cursor="")
                self._actualizar_estado("❌ Error al convertir a PDF", "red")
                return
 
            # 2️⃣ Enviar a Teams
            self._actualizar_estado("⏳ Enviando a Teams...", "#e67e22")
            self.root.update()
 
            try:
                exito = self.teams_webhook.enviar_archivo(pdf_file, email, mensaje)
                teams_win.config(cursor="")
                if exito:
                    self._actualizar_estado(
                        f"📤 PDF enviado a {email}: {os.path.basename(pdf_file)}", "#0078d4")
                    teams_win.destroy()
                else:
                    self._actualizar_estado("❌ Error al enviar a Teams", "red")
            except Exception as e:
                teams_win.config(cursor="")
                self._actualizar_estado("❌ Error al enviar", "red")
                messagebox.showerror("Error", f"Error al enviar:\n{str(e)}")
 
        tk.Button(btn_frame, text="🚀 Convertir y Enviar", command=enviar,
                  bg="#0078d4", fg="white", font=("Arial", 10, "bold"),
                  padx=30, pady=8, cursor="hand2").pack(side="left", padx=(0, 10))
        tk.Button(btn_frame, text="❌ Cancelar", command=teams_win.destroy,
                  bg="#666", fg="white", font=("Arial", 10),
                  padx=30, pady=8, cursor="hand2").pack(side="left")
 
 
# ============================================================
# 🚀 ENTRY POINT
# ============================================================
def main():
    root = tk.Tk()
    app  = FormularioApp(root)
    root.mainloop()
 
if __name__ == "__main__":
    main()
