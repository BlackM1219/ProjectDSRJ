# -*- coding: utf-8 -*-
"""
Formulario DISAGRO - Sistema completo con integración a Teams
+ Generador de Declaración Jurídica (Word)

INSTRUCCIONES:
1. pip install openpyxl pywin32 requests python-docx num2words
2. Configura tu WEBHOOK_URL
3. plantilla.xlsx y plantilla.docx en la misma carpeta
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from docx import Document
from docx.shared import Pt
import win32com.client
from typing import Optional, Dict
from dataclasses import dataclass, field
from num2words import num2words
import re
from datetime import datetime
import os
import requests
import base64

TEMPLATE_PATH      = "plantilla.xlsx"
TEMPLATE_DOCX_PATH = "plantilla.docx"

# ============================================
# 🔧 WEBHOOK URL
# ============================================
WEBHOOK_URL = "https://default93a2cd4474774ba69ca3e4e7aaf9c2.dd.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/0f4c9b26740b401c9681e227af854f4e/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=wfREfiqzfGGgVHBoQajhWsotqJs7m6_FbPD8zq1mnwc"


# ============================================================
# 🔄 CONVERSIÓN EXCEL → PDF
# ============================================================
def convertir_excel_a_pdf(excel_path: str) -> Optional[str]:
    excel_app = None
    wb = None
    try:
        excel_abs = os.path.abspath(excel_path)
        pdf_path  = excel_abs.replace('.xlsx', '.pdf')
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible       = False
        excel_app.DisplayAlerts = False
        wb = excel_app.Workbooks.Open(excel_abs)
        wb.ExportAsFixedFormat(0, pdf_path)
        return pdf_path
    except Exception as e:
        messagebox.showerror("❌ Error al convertir a PDF",
            f"No se pudo convertir el Excel a PDF.\n\nDetalle: {str(e)}")
        return None
    finally:
        try:
            if wb: wb.Close(False)
        except: pass
        try:
            if excel_app: excel_app.Quit()
        except: pass


# ============================================================
# 📝 UTILIDADES JURÍDICO
# ============================================================
def numero_a_letras(num):
    try:    return num2words(num, lang='es')
    except: return str(num)

def anio_a_letras(anio):
    if 2000 <= anio < 3000:
        resto = anio - 2000
        return "dos mil" if resto == 0 else f"dos mil {numero_a_letras(resto)}"
    return numero_a_letras(anio)

def dpi_a_letras(dpi_num):
    dpi_num = dpi_num.replace(" ","").replace(",","").strip()
    if not dpi_num.isdigit(): return dpi_num
    dpi_num = dpi_num.zfill(13)
    g1, g2, g3 = dpi_num[:4], dpi_num[4:9], dpi_num[9:]
    try:
        return (f"{numero_a_letras(int(g1)).capitalize()}, "
                f"{numero_a_letras(int(g2))}, "
                f"{numero_a_letras(int(g3))}")
    except: return dpi_num


# ============================================================
# 📦 DATOS
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
    def __init__(self, webhook_url):
        self.webhook_url = webhook_url

    def enviar_archivo(self, file_path, email_destinatario, mensaje="") -> bool:
        url = (self.webhook_url or "").strip()
        if not url or not url.startswith("http"):
            messagebox.showerror("⚠️ Webhook no configurado",
                "Configura el WEBHOOK_URL en el código.")
            return False
        try:
            with open(file_path, 'rb') as f:
                b64 = base64.b64encode(f.read()).decode('utf-8')
            payload = {
                "email":             email_destinatario,
                "mensaje":           mensaje or f"📋 {os.path.basename(file_path)}",
                "archivo_nombre":    os.path.basename(file_path),
                "archivo_contenido": b64
            }
            r = requests.post(url, json=payload,
                headers={"Content-Type": "application/json"}, timeout=30)
            if r.status_code in [200, 202, 204]:
                messagebox.showinfo("✅ Enviado",
                    f"Archivo enviado a {email_destinatario}")
                return True
            else:
                messagebox.showerror("❌ Error", f"Código {r.status_code}\n{r.text[:200]}")
        except Exception as e:
            messagebox.showerror("❌ Error", str(e))
        return False


# ============================================================
# 🔍 OUTLOOK
# ============================================================
class OutlookSearcher:
    @staticmethod
    def buscar_usuario(alias) -> Optional[Dict]:
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
                        parts = ex.Name.split()
                        fn = parts[0] if parts else ""
                        sn, fs, ss = "", "", ""
                        if len(parts) >= 3:
                            fs = parts[-2]; ss = parts[-1]
                            if len(parts) > 3:
                                sn = " ".join(parts[1:-2])
                        elif len(parts) == 2:
                            fs = parts[1]
                        return {
                            "alias": alias, "full_name": ex.Name,
                            "first_name": fn, "second_name": sn,
                            "first_surname": fs, "second_surname": ss,
                            "correo": correo, "puesto": ex.JobTitle or "",
                            "empresa": ex.CompanyName or "",
                            "depart": ex.Department or ""
                        }
                except AttributeError: continue
            return None
        except Exception as e:
            messagebox.showerror("Error COM", str(e))
            return None


# ============================================================
# 📊 EXCEL GENERATOR
# ============================================================
class ExcelGenerator:
    FORM_TYPES = [
        "FECHA DE DEVOLUCIÓN DE EQUIPO", "HOJA DE RESPONSABILIDAD",
        "SOPORTE TÉCNICO", "PASE DE SALIDA", "DICTAMEN TÉCNICO",
        "RECEPCIÓN DE EQUIPO", "PRÉSTAMO DE EQUIPO"
    ]

    def __init__(self, template_path=TEMPLATE_PATH):
        self.template_path = template_path

    def generar(self, data, output_dir=".") -> Optional[str]:
        if not data.ticket or not data.alias:
            messagebox.showwarning("Faltan datos",
                "Busca un usuario y completa el número de ticket.")
            return None
        try:
            wb = load_workbook(self.template_path)
            ws = wb.active
            if not data.fecha:
                data.fecha = datetime.now().strftime("%d/%m/%Y")
            self._reemplazar_etiquetas(ws, self._preparar_datos(data))
            self._marcar_checkboxes(ws, data)
            equipos = data.temp_equipos if data.temp_equipos else data.more_equipment
            if equipos: self._llenar_tabla_equipos(ws, equipos)
            sa = re.sub(r'[^\w\-]', '_', data.alias)
            st = re.sub(r'[^\w\-]', '_', data.ticket)
            out = os.path.join(output_dir, f"{st}_{sa}.xlsx")
            wb.save(out)
            return out
        except FileNotFoundError:
            messagebox.showerror("Error", f"No se encontró: {self.template_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        return None

    def _preparar_datos(self, data) -> Dict:
        return {
            "ticket": data.ticket, "fecha": data.fecha, "alias": data.alias,
            "first_name": data.first_name, "second_name": data.second_name,
            "first_surname": data.first_surname, "second_surname": data.second_surname,
            "puesto": data.puesto, "depart": data.depart, "empresa": data.empresa,
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

    def _reemplazar_etiquetas(self, ws, reps):
        pat = re.compile(r'\{\{(\w+)\}\}')
        for ri in range(1, ws.max_row + 1):
            for ci in range(1, ws.max_column + 1):
                try:
                    cell = ws.cell(row=ri, column=ci)
                    if isinstance(cell, MergedCell): continue
                    if cell.value and isinstance(cell.value, str) and "{{" in cell.value:
                        nv = cell.value
                        for tag in pat.findall(cell.value):
                            if tag in reps:
                                nv = nv.replace(f"{{{{{tag}}}}}", str(reps[tag]))
                        if nv != cell.value:
                            self._set(ws, ri, ci, nv)
                except: continue

    def _set(self, ws, row, col, value):
        try:
            c = ws.cell(row=row, column=col)
            if isinstance(c, MergedCell):
                for mr in ws.merged_cells.ranges:
                    if mr.min_row<=row<=mr.max_row and mr.min_col<=col<=mr.max_col:
                        ws.cell(row=mr.min_row, column=mr.min_col, value=value); return
            else: c.value = value
        except:
            try: ws.cell(row=row, column=col).value = value
            except: pass

    def _idx(self, letter):
        idx = 0
        for ch in letter.upper(): idx = idx*26+(ord(ch)-ord('A')+1)
        return idx

    def _marcar_checkboxes(self, ws, data):
        m = {
            "HOJA DE RESPONSABILIDAD":("B",6),"SOPORTE TÉCNICO":("M",6),
            "SOPORTE TECNICO":("M",6),"PASE DE SALIDA":("V",6),
            "DICTAMEN TÉCNICO":("B",7),"DICTAMEN TECNICO":("B",7),
            "RECEPCIÓN DE EQUIPO":("M",7),"RECEPCION DE EQUIPO":("M",7),
            "PRÉSTAMO DE EQUIPO":("V",7),"PRESTAMO DE EQUIPO":("V",7),
        }
        for ft in data.form_types:
            if not ft: continue
            k = ft.strip().upper().replace("Á","A").replace("É","E").replace("Í","I").replace("Ó","O").replace("Ú","U")
            if k in m:
                cl, r = m[k]
                try: self._set(ws, r, self._idx(cl), "X")
                except: pass
        if data.sistema_operativo:
            so = data.sistema_operativo.upper()
            if "WINDOWS" in so: self._set(ws,38,self._idx("J"),"X")
            if "MAC"     in so: self._set(ws,38,self._idx("N"),"X")
        if data.ofimática and ("OFFICE" in data.ofimática.upper() or "MS" in data.ofimática.upper()):
            self._set(ws,39,self._idx("J"),"X")

    def _llenar_tabla_equipos(self, ws, equipos):
        fi = 22
        cols = {"equipo":self._idx("C"),"marca":self._idx("M"),
                "modelo":self._idx("U"),"serie":self._idx("AC"),"estado":self._idx("AO")}
        for i, eq in enumerate(equipos):
            r = fi+i
            for k, c in cols.items():
                self._set(ws, r, c, eq.get(k,"" if k!="estado" else "OK"))


# ============================================================
# 🖥️  APP PRINCIPAL — layout con Notebook (pestañas)
# ============================================================
class FormularioApp:
    def __init__(self, root):
        self.root          = root
        self.root.title("Formulario DISAGRO - Departamento de Sistemas")
        self.root.geometry("900x680")
        self.root.resizable(True, True)
        self.root.columnconfigure(0, weight=1)

        self.data          = UserData()
        self.excel_gen     = ExcelGenerator()
        self.teams_webhook = TeamsWebhook(WEBHOOK_URL)
        self.widgets       = {}
        self.ultimo_excel  = None

        self._build_ui()

    # ----------------------------------------------------------
    # CONSTRUCCIÓN DE LA INTERFAZ
    # ----------------------------------------------------------
    def _build_ui(self):

        # ── Título ────────────────────────────────────────────
        hdr = tk.Frame(self.root, bg="#2c3e50", pady=8)
        hdr.grid(row=0, column=0, sticky="ew")
        tk.Label(hdr,
            text="SISTEMA DE FORMULARIOS  —  DEPARTAMENTO DE SISTEMAS",
            font=("Arial", 13, "bold"), bg="#2c3e50", fg="white").pack()

        # ── Búsqueda (siempre visible arriba) ─────────────────
        sf = tk.LabelFrame(self.root, text="Buscar Usuario", padx=10, pady=8)
        sf.grid(row=1, column=0, sticky="ew", padx=10, pady=(6,0))
        tk.Label(sf, text="Alias o correo:").grid(row=0, column=0, sticky="w")
        self.widgets['alias'] = tk.Entry(sf, width=38)
        self.widgets['alias'].grid(row=0, column=1, padx=5)
        self.widgets['alias'].bind('<Return>', lambda e: self._buscar())
        tk.Button(sf, text="🔍 Buscar en Outlook", command=self._buscar,
            bg="#3498db", fg="white", padx=10).grid(row=0, column=2)

        # Resultado búsqueda
        tf = tk.LabelFrame(self.root, text="Usuario encontrado", padx=8, pady=4)
        tf.grid(row=2, column=0, sticky="ew", padx=10, pady=(4,0))
        cols = ("Nombre","Correo","Puesto","Empresa","Departamento")
        self.widgets['tree'] = ttk.Treeview(tf, columns=cols, show="headings", height=2)
        for c in cols:
            self.widgets['tree'].heading(c, text=c)
            self.widgets['tree'].column(c, width=155)
        self.widgets['tree'].pack(fill="x")

        # ── Notebook con 2 pestañas ───────────────────────────
        nb = ttk.Notebook(self.root)
        nb.grid(row=3, column=0, sticky="nsew", padx=10, pady=8)
        self.root.rowconfigure(3, weight=1)

        # ── PESTAÑA 1: Excel / PDF / Teams ───────────────────
        tab1 = tk.Frame(nb, padx=20, pady=20)
        nb.add(tab1, text="  📊  Formulario Excel / PDF  ")

        # Indicador de usuario
        self.widgets['lbl_user_excel'] = tk.Label(
            tab1, text="⚠️  Busca un usuario para continuar",
            font=("Arial", 10, "italic"), fg="#c0392b")
        self.widgets['lbl_user_excel'].pack(pady=(0,12))

        self.widgets['btn_rellenar'] = tk.Button(
            tab1, text="📝  Rellenar Datos del Formulario",
            command=self._abrir_formulario_datos, state="disabled",
            bg="#e67e22", fg="white", padx=20, pady=12,
            font=("Arial", 11, "bold"), cursor="hand2", width=40)
        self.widgets['btn_rellenar'].pack(pady=4)

        fila_pdf = tk.Frame(tab1)
        fila_pdf.pack(pady=4)

        self.widgets['btn_excel'] = tk.Button(
            fila_pdf, text="✅  1. Generar Excel",
            command=self._generar_excel, state="disabled",
            bg="#27ae60", fg="white", padx=18, pady=12,
            font=("Arial", 11, "bold"), cursor="hand2")
        self.widgets['btn_excel'].grid(row=0, column=0, padx=6)

        tk.Label(fila_pdf, text="→", font=("Arial", 18)).grid(row=0, column=1)

        self.widgets['btn_enviar'] = tk.Button(
            fila_pdf, text="📤  2. Convertir PDF y Enviar a Teams",
            command=self._convertir_y_enviar, state="disabled",
            bg="#0078d4", fg="white", padx=18, pady=12,
            font=("Arial", 11, "bold"), cursor="hand2")
        self.widgets['btn_enviar'].grid(row=0, column=2, padx=6)

        self.widgets['lbl_estado'] = tk.Label(
            tab1, text="⬜  Sin archivo generado",
            font=("Arial", 9), fg="gray")
        self.widgets['lbl_estado'].pack(pady=(10,0))

        # ── PESTAÑA 2: Jurídico ──────────────────────────────
        tab2 = tk.Frame(nb, padx=20, pady=20)
        nb.add(tab2, text="  ⚖️  Declaración Jurídica  ")

        self.widgets['lbl_user_jur'] = tk.Label(
            tab2, text="⚠️  Busca un usuario para pre-cargar los datos",
            font=("Arial", 10, "italic"), fg="#c0392b")
        self.widgets['lbl_user_jur'].pack(pady=(0,10))

        tk.Label(tab2,
            text="Genera la Declaración de Responsabilidad en formato Word (.docx)\n"
                 "Los datos del colaborador se cargan automáticamente al buscar un usuario.",
            font=("Arial", 10), fg="#555", justify="center").pack(pady=(0,20))

        self.widgets['btn_juridico'] = tk.Button(
            tab2,
            text="⚖️   GENERAR DECLARACIÓN JURÍDICA",
            command=self._abrir_formulario_juridico,
            state="disabled",
            bg="#8e44ad", fg="white",
            padx=30, pady=18,
            font=("Arial", 13, "bold"),
            cursor="hand2", width=38,
            relief="raised", bd=3,
            activebackground="#6c3483",
            activeforeground="white")
        self.widgets['btn_juridico'].pack()

        tk.Label(tab2,
            text="Requiere: plantilla.docx en la misma carpeta",
            font=("Arial", 8), fg="gray").pack(pady=(12,0))

    # ----------------------------------------------------------
    def _actualizar_estado(self, texto, color="gray"):
        self.widgets['lbl_estado'].config(text=texto, fg=color)

    # ----------------------------------------------------------
    # BÚSQUEDA
    # ----------------------------------------------------------
    def _buscar(self):
        alias = self.widgets['alias'].get().strip()
        if not alias:
            messagebox.showwarning("Campo vacío", "Ingresa un alias o correo.")
            return
        self.widgets['tree'].delete(*self.widgets['tree'].get_children())
        res = OutlookSearcher.buscar_usuario(alias)
        if res:
            for k, v in res.items(): setattr(self.data, k, v)
            self.widgets['tree'].insert("", "end", values=(
                res["full_name"], res["correo"],
                res["puesto"], res["empresa"], res["depart"]))

            # Habilitar todos los botones
            for btn in ('btn_rellenar','btn_excel','btn_juridico'):
                self.widgets[btn].config(state="normal")

            # Actualizar indicadores de usuario
            msg = f"✅  Usuario cargado: {res['full_name']}"
            self.widgets['lbl_user_excel'].config(text=msg, fg="#27ae60")
            self.widgets['lbl_user_jur'].config(text=msg, fg="#27ae60")

            messagebox.showinfo("✓ Usuario encontrado",
                f"Usuario {res['full_name']} cargado correctamente.")
        else:
            for btn in ('btn_rellenar','btn_excel','btn_juridico'):
                self.widgets[btn].config(state="disabled")
            msg = "⚠️  Busca un usuario para continuar"
            self.widgets['lbl_user_excel'].config(text=msg, fg="#c0392b")
            self.widgets['lbl_user_jur'].config(text=msg, fg="#c0392b")
            messagebox.showwarning("No encontrado", f"No se encontró: {alias}")

    # ----------------------------------------------------------
    # FORMULARIO DE DATOS (Excel)
    # ----------------------------------------------------------
    def _abrir_formulario_datos(self):
        if not self.data.alias:
            messagebox.showwarning("Busca primero", "Primero busca un usuario.")
            return

        win = tk.Toplevel(self.root)
        win.title(f"Datos — {self.data.full_name}")
        win.geometry("800x880")

        canvas = tk.Canvas(win)
        sb     = tk.Scrollbar(win, orient="vertical", command=canvas.yview)
        fr     = tk.Frame(canvas)
        fr.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=fr, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        entries = {}
        row = 0

        def sec(text, color="#34495e"):
            nonlocal row
            tk.Label(fr, text=text, font=("Arial",11,"bold"),
                     bg=color, fg="white").grid(
                row=row, column=0, columnspan=2, sticky="ew", pady=(12,4))
            row += 1

        def field(label, key, default="", widget=None):
            nonlocal row
            tk.Label(fr, text=label).grid(row=row, column=0, sticky="e", padx=5, pady=4)
            if widget is None:
                w = tk.Entry(fr, width=50)
                w.grid(row=row, column=1, sticky="w", padx=5, pady=4)
                w.insert(0, default)
            else:
                w = widget
                w.grid(row=row, column=1, sticky="w", padx=5, pady=4)
            entries[key] = w
            row += 1

        # Encabezado
        tk.Label(fr, text="DATOS DEL FORMULARIO", font=("Arial",12,"bold"),
                 bg="#2c3e50", fg="white").grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=(0,8))
        row += 1

        field("No. de Ticket *", "ticket", self.data.ticket)
        field("Fecha",           "fecha",  datetime.now().strftime("%d/%m/%Y"))

        # Tipo de formulario
        sec("TIPO DE FORMULARIO")
        tk.Label(fr, text="Selecciona tipo(s) *").grid(
            row=row, column=0, sticky="ne", padx=5, pady=4)
        ff = tk.Frame(fr)
        ff.grid(row=row, column=1, sticky="w", padx=5, pady=4)
        form_vars = {}
        for i, ft in enumerate(ExcelGenerator.FORM_TYPES):
            var = tk.BooleanVar(value=ft in self.data.form_types)
            tk.Checkbutton(ff, text=ft, variable=var).grid(row=i, column=0, sticky="w")
            form_vars[ft] = var
        entries['form_types'] = form_vars
        row += 1

        # Sección A
        sec("A. DATOS DEL COLABORADOR (cargados desde Outlook)")
        field("Localidad", "localidad", self.data.localidad)
        field("País",      "pais",      self.data.pais)

        # Sección B
        sec("B. HARDWARE")
        campos_hw = [
            ("Tipo de Computadora *","pc_type"),("Marca *","marca"),
            ("Modelo","model"),("Service Tag (ST) *","st"),
            ("Fecha de Compra","fecha_compra"),("Disco Duro","disco_duro"),
            ("Memoria RAM","memoria_ram"),("IP","ip"),("Host Name","hostname"),
        ]
        for lbl, key in campos_hw:
            field(lbl, key, getattr(self.data, key, ""))

        tk.Label(fr, text="Estado del equipo").grid(
            row=row, column=0, sticky="e", padx=5, pady=4)
        cb_estado = ttk.Combobox(fr, values=["OK","MALO","REGULAR"],
                                  state="readonly", width=47)
        cb_estado.set(self.data.estado)
        cb_estado.grid(row=row, column=1, sticky="w", padx=5, pady=4)
        entries['estado'] = cb_estado
        row += 1

        # Sección C
        sec("C. EQUIPOS ADICIONALES")
        eq_frame = tk.Frame(fr, relief="sunken", bd=1)
        eq_frame.grid(row=row, column=0, columnspan=2, padx=5, pady=4, sticky="ew")
        eq_list = tk.Listbox(eq_frame, height=4, width=80)
        eq_list.pack(side="left", fill="both", expand=True)
        eq_sb2 = tk.Scrollbar(eq_frame, command=eq_list.yview)
        eq_sb2.pack(side="right", fill="y")
        eq_list.config(yscrollcommand=eq_sb2.set)
        for eq in self.data.more_equipment:
            eq_list.insert("end", f"{eq.get('equipo','')} - {eq.get('marca','')} {eq.get('modelo','')} - {eq.get('serie','')}")
        row += 1

        btn_eq = tk.Frame(fr)
        btn_eq.grid(row=row, column=0, columnspan=2)

        def agregar_equipo():
            ew = tk.Toplevel(win); ew.title("Agregar Equipo"); ew.geometry("440x280")
            ee = {}
            for i,(lbl,k) in enumerate(zip(
                ["Equipo:","Marca:","Modelo:","Serie:","Estado:"],
                ["equipo","marca","modelo","serie","estado"]
            )):
                tk.Label(ew, text=lbl).grid(row=i,column=0,sticky="e",padx=5,pady=4)
                if k=="estado":
                    w=ttk.Combobox(ew,values=["OK","MALO","REGULAR"],state="readonly",width=28); w.set("OK")
                else:
                    w=tk.Entry(ew,width=30)
                w.grid(row=i,column=1,padx=5,pady=4); ee[k]=w
            def save():
                d={k:v.get().strip() for k,v in ee.items()}
                if d["equipo"]:
                    eq_list.insert("end",f"{d['equipo']} - {d['marca']} {d['modelo']} - {d['serie']}")
                    self.data.temp_equipos.append(d); ew.destroy()
                else: messagebox.showwarning("","Nombre del equipo obligatorio")
            tk.Button(ew,text="💾 Guardar",command=save,bg="#27ae60",fg="white").grid(row=5,column=0,columnspan=2,pady=8)

        def eliminar_equipo():
            sel = eq_list.curselection()
            if not sel: return
            idx=sel[0]; eq_list.delete(idx)
            if idx<len(self.data.temp_equipos): self.data.temp_equipos.pop(idx)

        tk.Button(btn_eq,text="➕ Agregar", command=agregar_equipo,bg="#3498db",fg="white").grid(row=0,column=0,padx=5,pady=4)
        tk.Button(btn_eq,text="➖ Eliminar",command=eliminar_equipo,bg="#c0392b",fg="white").grid(row=0,column=1,padx=5,pady=4)
        row += 1

        # Sección D
        sec("D. SOFTWARE")
        tk.Label(fr, text="Sistema operativo").grid(row=row,column=0,sticky="ne",padx=5,pady=4)
        so_f=tk.Frame(fr); so_f.grid(row=row,column=1,sticky="w",padx=5,pady=4)
        so_vars={}
        for i,opt in enumerate(["Windows","Mac OS"]):
            var=tk.BooleanVar(value=opt in self.data.sistema_operativo)
            tk.Checkbutton(so_f,text=opt,variable=var).grid(row=0,column=i,sticky="w",padx=8)
            so_vars[opt]=var
        entries['so_vars']=so_vars; row+=1

        field("Versión SO",    "version_so",    self.data.version_so)

        tk.Label(fr,text="Ofimática").grid(row=row,column=0,sticky="ne",padx=5,pady=4)
        of_f=tk.Frame(fr); of_f.grid(row=row,column=1,sticky="w",padx=5,pady=4)
        of_vars={}
        for i,opt in enumerate(["MS Office","Open Office"]):
            var=tk.BooleanVar(value=opt in self.data.ofimática)
            tk.Checkbutton(of_f,text=opt,variable=var).grid(row=0,column=i,sticky="w",padx=8)
            of_vars[opt]=var
        entries['of_vars']=of_vars; row+=1

        field("Versión Office","version_office",self.data.version_office)

        # Sección E/F
        sec("E. PROBLEMA / F. DIAGNÓSTICO")
        for lbl,key,h in [
            ("Problema reportado:","reported_problem",3),
            ("Diagnóstico:","diagnosis",5),
            ("Documentación/Rep.:","documentation",2),
        ]:
            tk.Label(fr,text=lbl).grid(row=row,column=0,sticky="ne",padx=5,pady=4)
            t=tk.Text(fr,width=50,height=h)
            t.grid(row=row,column=1,sticky="w",padx=5,pady=4)
            v=getattr(self.data,key,"")
            if v: t.insert("1.0",v)
            entries[key]=t; row+=1

        # Guardar
        ttk.Separator(fr,orient="horizontal").grid(row=row,column=0,columnspan=2,sticky="ew",pady=(12,4))
        row+=1

        def guardar():
            if not entries['ticket'].get().strip():
                messagebox.showwarning("","Ticket obligatorio"); return False
            self.data.ticket     = entries['ticket'].get().strip()
            self.data.fecha      = entries['fecha'].get().strip()
            self.data.form_types = [k for k,v in entries['form_types'].items() if v.get()]
            self.data.localidad  = entries['localidad'].get().strip()
            self.data.pais       = entries['pais'].get().strip()
            for _,key in campos_hw: setattr(self.data,key,entries[key].get().strip())
            self.data.estado = entries['estado'].get().strip()
            self.data.sistema_operativo = ", ".join(s for s,v in entries['so_vars'].items() if v.get())
            self.data.version_so    = entries['version_so'].get().strip()
            self.data.ofimática     = ", ".join(o for o,v in entries['of_vars'].items() if v.get())
            self.data.version_office= entries['version_office'].get().strip()
            self.data.reported_problem=entries['reported_problem'].get("1.0","end").strip()
            self.data.diagnosis       =entries['diagnosis'].get("1.0","end").strip()
            self.data.documentation   =entries['documentation'].get("1.0","end").strip()
            return True

        bf = tk.Frame(fr); bf.grid(row=row,column=0,columnspan=2,pady=14)
        tk.Button(bf,text="💾 Guardar",command=lambda:[guardar() and (messagebox.showinfo("","Guardado.") or win.destroy())],
            bg="#27ae60",fg="white",padx=16,pady=8,font=("Arial",10,"bold"),cursor="hand2").pack(side="left",padx=6)
        tk.Button(bf,text="⚖️ Guardar y abrir Jurídico",
            command=lambda:[guardar() and (win.destroy() or self._abrir_formulario_juridico())],
            bg="#8e44ad",fg="white",padx=16,pady=8,font=("Arial",10,"bold"),cursor="hand2").pack(side="left",padx=6)

    # ----------------------------------------------------------
    # PASO 1: GENERAR EXCEL
    # ----------------------------------------------------------
    def _generar_excel(self):
        if not self.data.alias:
            messagebox.showwarning("","Busca un usuario primero."); return
        if not self.data.ticket:
            messagebox.showwarning("","Completa el ticket en 'Rellenar Datos'."); return
        d = filedialog.askdirectory(title="Carpeta para guardar", initialdir=os.getcwd())
        if not d:
            if not messagebox.askyesno("","¿Guardar en la carpeta actual?"): return
            d = os.getcwd()
        f = self.excel_gen.generar(self.data, output_dir=d)
        if f:
            self.ultimo_excel = f
            self.widgets['btn_enviar'].config(state="normal")
            self._actualizar_estado(f"✅ Excel listo: {os.path.basename(f)}", "#27ae60")
            self.data.temp_equipos = []
            messagebox.showinfo("✅ Excel guardado", f)

    # ----------------------------------------------------------
    # PASO 2: CONVERTIR A PDF Y ENVIAR
    # ----------------------------------------------------------
    def _convertir_y_enviar(self):
        if not self.ultimo_excel or not os.path.exists(self.ultimo_excel):
            messagebox.showwarning("","Genera el Excel primero."); return

        tw = tk.Toplevel(self.root)
        tw.title("Convertir y Enviar"); tw.geometry("500x360"); tw.resizable(False,False)
        mf = tk.Frame(tw,padx=20,pady=20); mf.pack(fill="both",expand=True)

        tk.Label(mf,text="📤 Convertir PDF y Enviar a Teams",font=("Arial",13,"bold"),fg="#0078d4").pack(pady=(0,12))
        tk.Label(mf,text="Archivo:",font=("Arial",10,"bold")).pack(anchor="w")
        tk.Label(mf,text=os.path.basename(self.ultimo_excel),font=("Arial",9),fg="gray").pack(anchor="w",pady=(0,10))
        tk.Label(mf,text="Email del destinatario:",font=("Arial",10,"bold")).pack(anchor="w")
        ee = tk.Entry(mf,width=50,font=("Arial",10)); ee.pack(fill="x",pady=(4,10))
        if self.data.correo: ee.insert(0,self.data.correo)
        tk.Label(mf,text="Mensaje (opcional):",font=("Arial",10,"bold")).pack(anchor="w")
        mt = tk.Text(mf,width=50,height=3,font=("Arial",9)); mt.pack(fill="x",pady=(4,10))
        mt.insert("1.0",f"📋 Formulario de {self.data.full_name} - Ticket #{self.data.ticket}")

        bf = tk.Frame(mf); bf.pack(fill="x",pady=(8,0))

        def enviar():
            email=ee.get().strip()
            if not email or "@" not in email:
                messagebox.showwarning("","Email inválido."); return
            tw.config(cursor="wait"); tw.update()
            self._actualizar_estado("⏳ Convirtiendo a PDF…","#e67e22"); self.root.update()
            pdf = convertir_excel_a_pdf(self.ultimo_excel)
            if not pdf or not os.path.exists(pdf):
                tw.config(cursor=""); self._actualizar_estado("❌ Error al convertir","red"); return
            self._actualizar_estado("⏳ Enviando a Teams…","#e67e22"); self.root.update()
            try:
                ok = self.teams_webhook.enviar_archivo(pdf, email, mt.get("1.0","end").strip())
                tw.config(cursor="")
                if ok:
                    self._actualizar_estado(f"📤 Enviado a {email}","#0078d4"); tw.destroy()
                else:
                    self._actualizar_estado("❌ Error al enviar","red")
            except Exception as e:
                tw.config(cursor=""); self._actualizar_estado("❌ Error","red")
                messagebox.showerror("Error",str(e))

        tk.Button(bf,text="🚀 Convertir y Enviar",command=enviar,
            bg="#0078d4",fg="white",font=("Arial",10,"bold"),
            padx=28,pady=8,cursor="hand2").pack(side="left",padx=(0,8))
        tk.Button(bf,text="❌ Cancelar",command=tw.destroy,
            bg="#666",fg="white",font=("Arial",10),
            padx=28,pady=8,cursor="hand2").pack(side="left")

    # ----------------------------------------------------------
    # ⚖️ DECLARACIÓN JURÍDICA
    # ----------------------------------------------------------
    def _abrir_formulario_juridico(self):
        win = tk.Toplevel(self.root)
        win.title("Generador de Declaración de Responsabilidad")
        win.geometry("620x580")
        win.resizable(True, True)

        # Título
        tk.Label(win, text="⚖️  GENERADOR DE DECLARACIÓN JURÍDICA",
            font=("Arial", 12, "bold"), bg="#8e44ad", fg="white").pack(
            fill="x", ipady=8)

        # Canvas con scroll para todos los campos
        canvas  = tk.Canvas(win)
        scrollb = tk.Scrollbar(win, orient="vertical", command=canvas.yview)
        sf      = tk.Frame(canvas, padx=15, pady=10)
        sf.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=sf, anchor="nw")
        canvas.configure(yscrollcommand=scrollb.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollb.pack(side="right", fill="y")

        # Scroll con rueda del ratón
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        entries = {}

        def add_field(label, key, default="", width=45):
            row = len(entries)
            tk.Label(sf, text=f"{label}:", font=("Arial", 9),
                     anchor="e", width=22).grid(
                row=row, column=0, padx=5, pady=4, sticky="e")
            e = tk.Entry(sf, width=width, font=("Arial", 9))
            e.grid(row=row, column=1, padx=5, pady=4, sticky="w")
            if default:
                e.insert(0, default)
            entries[key] = e

        # ── Sección: Datos del colaborador ────────────────────
        tk.Label(sf, text="DATOS DEL COLABORADOR",
                 font=("Arial", 10, "bold"), bg="#8e44ad", fg="white").grid(
            row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6), ipady=3)

        # Reemplazar row tracking — usamos lista para poder mutar en closure
        rows = [1]

        def add_row(label, key, default="", width=45):
            r = rows[0]
            tk.Label(sf, text=f"{label}:", font=("Arial", 9),
                     anchor="e", width=22).grid(
                row=r, column=0, padx=5, pady=4, sticky="e")
            e = tk.Entry(sf, width=width, font=("Arial", 9))
            e.grid(row=r, column=1, padx=5, pady=4, sticky="w")
            if default:
                e.insert(0, default)
            entries[key] = e
            rows[0] += 1

        # Datos auto-cargados desde Outlook (editables)
        add_row("Nombre completo",  "nombre",   self.data.full_name)
        add_row("Empresa",          "empresa",  self.data.empresa)

        # Separador
        tk.Frame(sf, bg="#ddd", height=1).grid(
            row=rows[0], column=0, columnspan=2, sticky="ew", pady=4)
        rows[0] += 1
        tk.Label(sf, text="DATOS PERSONALES",
                 font=("Arial", 10, "bold"), bg="#34495e", fg="white").grid(
            row=rows[0], column=0, columnspan=2, sticky="ew", pady=(0, 4), ipady=3)
        rows[0] += 1

        add_row("Edad",                  "edad",          "")
        add_row("Estado Civil",          "estado_civil",  "")
        add_row("Nacionalidad",          "nacionalidad",  "guatemalteco/a")
        add_row("Profesión / Puesto",    "profesion",     self.data.puesto)
        add_row("Domicilio",             "domicilio",     "")
        add_row("DPI (números)",         "dpi_numeros",   "")
        add_row("DPI (letras, opcional)","dpi_letras",    "")

        # Separador
        tk.Frame(sf, bg="#ddd", height=1).grid(
            row=rows[0], column=0, columnspan=2, sticky="ew", pady=4)
        rows[0] += 1
        tk.Label(sf, text="DATOS DEL EQUIPO  (cargados del formulario Excel)",
                 font=("Arial", 10, "bold"), bg="#34495e", fg="white").grid(
            row=rows[0], column=0, columnspan=2, sticky="ew", pady=(0, 4), ipady=3)
        rows[0] += 1

        # Marca, Modelo y ST vienen del formulario Excel
        add_row("Marca del equipo",  "marca",       self.data.marca)
        add_row("Modelo del equipo", "modelo",       self.data.model)
        add_row("Service Tag",       "service_tag",  self.data.st)

        # Separador
        tk.Frame(sf, bg="#ddd", height=1).grid(
            row=rows[0], column=0, columnspan=2, sticky="ew", pady=4)
        rows[0] += 1
        tk.Label(sf, text="ARCHIVO DE SALIDA",
                 font=("Arial", 10, "bold"), bg="#34495e", fg="white").grid(
            row=rows[0], column=0, columnspan=2, sticky="ew", pady=(0, 4), ipady=3)
        rows[0] += 1

        nombre_archivo_default = (
            f"Declaracion_{self.data.first_surname or 'Apellido'}"
            f"_{self.data.first_name or 'Nombre'}"
        )
        add_row("Nombre del archivo", "nombre_archivo", nombre_archivo_default)

        # ── Botones ───────────────────────────────────────────
        bf = tk.Frame(sf)
        bf.grid(row=rows[0], column=0, columnspan=2, pady=18)

        def limpiar():
            for key in ["edad","estado_civil","domicilio","dpi_numeros","dpi_letras"]:
                entries[key].delete(0, tk.END)

        def generar():
            if not os.path.exists(TEMPLATE_DOCX_PATH):
                messagebox.showerror("Error",
                    f"No se encontró '{TEMPLATE_DOCX_PATH}'.\n"
                    "Asegúrate de que esté en la misma carpeta.")
                return

            nombre  = entries['nombre'].get().strip()
            empresa = entries['empresa'].get().strip()
            if not nombre:
                messagebox.showwarning("", "El nombre es obligatorio."); return

            meses = ["enero","febrero","marzo","abril","mayo","junio",
                     "julio","agosto","septiembre","octubre","noviembre","diciembre"]
            hoy = datetime.now()

            # DPI letras: automático o manual
            dpi_num = entries['dpi_numeros'].get().replace(" ","").replace(",","")
            try:    dpi_l = dpi_a_letras(dpi_num)
            except: dpi_l = entries['dpi_letras'].get()
            if not dpi_l:
                dpi_l = entries['dpi_letras'].get()

            # Edad en letras
            try:    edad_l = numero_a_letras(int(entries['edad'].get()))
            except: edad_l = entries['edad'].get()

            datos = {
                "{NOMBRE}":       nombre,
                "{EDAD}":         edad_l,
                "{ESTADO_CIVIL}": entries['estado_civil'].get().strip(),
                "{NACIONALIDAD}": entries['nacionalidad'].get().strip(),
                "{PROFESION}":    entries['profesion'].get().strip(),
                "{DOMICILIO}":    entries['domicilio'].get().strip(),
                "{DPI_NUMEROS}":  entries['dpi_numeros'].get().strip(),
                "{DPI_LETRAS}":   dpi_l,
                "{MARCA}":        entries['marca'].get().strip(),
                "{MODELO}":       entries['modelo'].get().strip(),
                "{SERVICE_TAG}":  entries['service_tag'].get().strip(),
                "{EMPRESA}":      empresa,
                "{DIA}":          numero_a_letras(hoy.day),
                "{MES}":          meses[hoy.month - 1],
                "{ANIO}":         anio_a_letras(hoy.year),
            }

            print(f"DEBUG Jurídico → {hoy.day} de {meses[hoy.month-1]} de {anio_a_letras(hoy.year)}")

            try:
                doc = Document(TEMPLATE_DOCX_PATH)

                def rr(p):
                    for run in p.runs:
                        for k, v in datos.items():
                            if k in run.text:
                                run.text = run.text.replace(k, str(v))

                for p in doc.paragraphs:
                    rr(p)
                for table in doc.tables:
                    for trow in table.rows:
                        for cell in trow.cells:
                            for p in cell.paragraphs:
                                rr(p)

                # Unir párrafos partidos
                for i in range(len(doc.paragraphs) - 1):
                    if (doc.paragraphs[i].text.strip().endswith("causado") and
                            doc.paragraphs[i+1].text.strip().startswith("por no cumplir")):
                        doc.paragraphs[i].text = (doc.paragraphs[i].text.strip()
                                                  + " " + doc.paragraphs[i+1].text.strip())
                        doc.paragraphs[i+1].text = ""

                # Fuente Century Gothic
                for p in doc.paragraphs:
                    for run in p.runs:
                        if run.text.strip():
                            run.font.name = "Century Gothic"
                            run.font.size = Pt(11)
                for table in doc.tables:
                    for trow in table.rows:
                        for cell in trow.cells:
                            for p in cell.paragraphs:
                                for run in p.runs:
                                    if run.text.strip():
                                        run.font.name = "Century Gothic"
                                        run.font.size = Pt(11)

                nf = entries['nombre_archivo'].get().strip()
                if not nf:
                    nf = f"Declaracion_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                if not nf.lower().endswith(".docx"):
                    nf += ".docx"

                doc.save(nf)

                msg = (f"✅ Documento generado exitosamente:\n\n"
                       f"📁 {nf}\n\n"
                       f"📅 Fecha: {numero_a_letras(hoy.day)} de "
                       f"{meses[hoy.month-1]} de {anio_a_letras(hoy.year)}\n"
                       f"👤 Nombre: {nombre}\n"
                       f"🆔 DPI: {entries['dpi_numeros'].get()}")
                messagebox.showinfo("Documento Generado", msg)

                if messagebox.askyesno("Abrir archivo",
                        "¿Desea abrir el documento generado?"):
                    os.startfile(nf)

                canvas.unbind_all("<MouseWheel>")
                win.destroy()

            except Exception as e:
                messagebox.showerror("Error",
                    f"Ocurrió un error al generar el documento:\n\n{str(e)}")

        tk.Button(bf, text="🔄 Generar Documento",
                  command=generar,
                  bg="#27ae60", fg="white", font=("Arial", 10, "bold"),
                  width=18, height=2, cursor="hand2").pack(side="left", padx=5)

        tk.Button(bf, text="🗑️ Limpiar",
                  command=limpiar,
                  bg="#e74c3c", fg="white", font=("Arial", 10, "bold"),
                  width=12, height=2, cursor="hand2").pack(side="left", padx=5)

        tk.Button(bf, text="❌ Cancelar",
                  command=lambda: [canvas.unbind_all("<MouseWheel>"), win.destroy()],
                  bg="#666", fg="white", font=("Arial", 10, "bold"),
                  width=10, height=2, cursor="hand2").pack(side="left", padx=5)

        # Info
        tk.Label(sf,
            text="ℹ️  Marca, Modelo y Service Tag se cargan del formulario Excel.\n"
                 "    Asegúrate de que 'plantilla.docx' esté en la misma carpeta.",
            fg="gray", font=("Arial", 8), justify="left").grid(
            row=rows[0]+1, column=0, columnspan=2, pady=(0, 10))


# ============================================================
# 🚀 ENTRY POINT
# ============================================================
def main():
    root = tk.Tk()
    FormularioApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()