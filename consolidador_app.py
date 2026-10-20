"""Sistema Consolidador de Tramas - interfaz grafica (Tkinter)."""
import os
import queue
import sys
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, simpledialog, ttk

import consolidador_core as core

APP_NOMBRE = "Sistema Consolidador de Tramas"
APP_VERSION = "1.0"


def fmt_fecha(iso):
    if not iso:
        return "-"
    try:
        return datetime.fromisoformat(iso).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return iso


def fmt_bytes(n):
    for u in ("B", "KB", "MB", "GB"):
        if n < 1024 or u == "GB":
            return f"{n:.0f} {u}" if u == "B" else f"{n:.1f} {u}"
        n /= 1024


def abrir_carpeta(ruta):
    os.makedirs(ruta, exist_ok=True)
    if sys.platform.startswith("win"):
        os.startfile(ruta)
    else:
        import subprocess
        subprocess.Popen(["xdg-open", ruta])


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NOMBRE} v{APP_VERSION}")
        self.geometry("1000x680")
        self.minsize(860, 560)
        self.gestor = core.Gestor()
        self.cola = queue.Queue()
        self.ocupado = False
        self.cancelar = threading.Event()
        self.nombres_custom = {}   # ruta trama -> nombre de salida editado
        self.botones = []

        self._crear_menu()
        self._crear_ui()
        self.after(100, self._procesar_cola)
        self.refrescar_maestros()
        self.refrescar_tramas()
        # Sincronizacion inicial (importa Excel legado / registra maestros existentes)
        self.ejecutar(lambda: self.gestor.sincronizar(self.log), self.refrescar_maestros, "Revisando maestros...")

    # ------------------------------------------------------------------ UI
    def _crear_menu(self):
        barra = tk.Menu(self)
        m_arch = tk.Menu(barra, tearoff=0)
        m_arch.add_command(label="Abrir carpeta de salida", command=lambda: abrir_carpeta(self.gestor.rutas.salida))
        m_arch.add_command(label="Abrir carpeta de tramas", command=lambda: abrir_carpeta(self.gestor.rutas.tramas))
        m_arch.add_command(label="Abrir carpeta de maestros",
                           command=lambda: abrir_carpeta(self.gestor.rutas.maestros))
        m_arch.add_separator()
        m_arch.add_command(label="Salir", command=self.destroy)
        barra.add_cascade(label="Archivo", menu=m_arch)

        m_mae = tk.Menu(barra, tearoff=0)
        m_mae.add_command(label="Ver datos del maestro seleccionado", command=self.ver_maestro)
        m_mae.add_separator()
        m_act = tk.Menu(m_mae, tearoff=0)
        for clave, spec in core.MAESTROS.items():
            m_act.add_command(label=spec["nombre"], command=lambda c=clave: self.actualizar_maestro(c))
        m_mae.add_cascade(label="Actualizar maestro", menu=m_act)
        m_mae.add_command(label="Actualizar tablas desde Excel completo...", command=self.actualizar_excel_tablas)
        m_mae.add_separator()
        m_mae.add_command(label="Refrescar", command=self.refrescar_maestros)
        barra.add_cascade(label="Datos maestros", menu=m_mae)

        m_tr = tk.Menu(barra, tearoff=0)
        m_tr.add_command(label="Agregar tramas (CSV / ZIP)...", command=self.agregar_tramas)
        m_tr.add_command(label="Generar consolidado(s)", command=self.generar)
        barra.add_cascade(label="Tramas", menu=m_tr)

        m_ayu = tk.Menu(barra, tearoff=0)
        m_ayu.add_command(label="Instrucciones", command=self.instrucciones)
        m_ayu.add_command(label="Acerca de", command=lambda: messagebox.showinfo(
            "Acerca de", f"{APP_NOMBRE}\nVersion {APP_VERSION}\n\nGenera consolidados de tramas homologadas "
                         "con los datos maestros.\nLicencia MIT."))
        barra.add_cascade(label="Ayuda", menu=m_ayu)
        self.config(menu=barra)

    def _crear_ui(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=(8, 4))
        self._tab_maestros()
        self._tab_tramas()

        pie = ttk.Frame(self)
        pie.pack(fill="x", padx=8, pady=(0, 4))
        self.barra = ttk.Progressbar(pie, mode="determinate", maximum=100)
        self.barra.pack(side="left", fill="x", expand=True)
        self.lbl_estado = ttk.Label(pie, text="Listo", width=45, anchor="e")
        self.lbl_estado.pack(side="left", padx=8)

        marco_log = ttk.LabelFrame(self, text="Registro")
        marco_log.pack(fill="both", padx=8, pady=(0, 8))
        self.txt = tk.Text(marco_log, height=8, state="disabled", wrap="word")
        sb = ttk.Scrollbar(marco_log, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.txt.pack(fill="both", expand=True)

    def _boton(self, padre, texto, cmd, **pack):
        b = ttk.Button(padre, text=texto, command=cmd)
        b.pack(**pack)
        self.botones.append(b)
        return b

    def _tab_maestros(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text="Datos maestros")
        cols = ("nombre", "estado", "registros", "fecha", "origen")
        self.tree_m = ttk.Treeview(f, columns=cols, show="headings", selectmode="browse", height=9)
        for c, t, w in (("nombre", "Maestro", 190), ("estado", "Estado", 90), ("registros", "Registros", 100),
                        ("fecha", "Ultima actualizacion", 150), ("origen", "Archivo de origen", 420)):
            self.tree_m.heading(c, text=t)
            self.tree_m.column(c, width=w, anchor="w")
        self.tree_m.pack(fill="both", expand=True, padx=8, pady=8)
        self.tree_m.bind("<Double-1>", lambda e: self.ver_maestro())
        bar = ttk.Frame(f)
        bar.pack(fill="x", padx=8, pady=(0, 8))
        self._boton(bar, "Ver datos", self.ver_maestro, side="left", padx=2)
        self._boton(bar, "Actualizar seleccionado (CSV / ZIP / Excel)...", self.actualizar_seleccionado,
                    side="left", padx=2)
        self._boton(bar, "Tablas desde Excel completo...", self.actualizar_excel_tablas, side="left", padx=2)

    def _tab_tramas(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text="Tramas y consolidados")
        bar = ttk.Frame(f)
        bar.pack(fill="x", padx=8, pady=8)
        self._boton(bar, "Agregar tramas (CSV / ZIP)...", self.agregar_tramas, side="left", padx=2)
        self._boton(bar, "Quitar seleccionadas", self.quitar_tramas, side="left", padx=2)
        self._boton(bar, "Seleccionar todas", lambda: self.tree_t.selection_set(self.tree_t.get_children()),
                    side="left", padx=2)
        self._boton(bar, "Restablecer nombres", self.restablecer_nombres, side="left", padx=2)

        cols = ("trama", "tam", "salida")
        self.tree_t = ttk.Treeview(f, columns=cols, show="headings", selectmode="extended", height=9)
        for c, t, w in (("trama", "Trama", 380), ("tam", "Tamano", 90), ("salida", "Nombre del consolidado "
                                                                                "(doble clic para editar)", 400)):
            self.tree_t.heading(c, text=t)
            self.tree_t.column(c, width=w, anchor="w")
        self.tree_t.pack(fill="both", expand=True, padx=8)
        self.tree_t.bind("<Double-1>", self.editar_nombre)

        op = ttk.LabelFrame(f, text="Modo de salida")
        op.pack(fill="x", padx=8, pady=8)
        self.modo = tk.StringVar(value="separados")
        ttk.Radiobutton(op, text="Consolidados separados (uno por trama, con el nombre indicado arriba)",
                        variable=self.modo, value="separados", command=self._modo_cambio).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=6, pady=2)
        ttk.Radiobutton(op, text="Un solo consolidado (une todas las tramas seleccionadas, sin duplicados)",
                        variable=self.modo, value="unico", command=self._modo_cambio).grid(
            row=1, column=0, columnspan=3, sticky="w", padx=6, pady=2)
        ttk.Label(op, text="Nombre del consolidado unico:").grid(row=2, column=0, sticky="w", padx=(26, 4), pady=4)
        self.nombre_unico = tk.StringVar(value="CONSOLIDADO_COMPLETO")
        self.ent_unico = ttk.Entry(op, textvariable=self.nombre_unico, width=40, state="disabled")
        self.ent_unico.grid(row=2, column=1, sticky="w")
        self.bom = tk.BooleanVar(value=False)
        ttk.Checkbutton(op, text="Compatible con Excel (UTF-8 con BOM, muestra bien las tildes)",
                        variable=self.bom).grid(row=3, column=0, columnspan=3, sticky="w", padx=6, pady=2)

        bot = ttk.Frame(f)
        bot.pack(fill="x", padx=8, pady=(0, 8))
        self.btn_gen = self._boton(bot, "Generar consolidado(s)", self.generar, side="left", padx=2)
        self.btn_cancel = ttk.Button(bot, text="Cancelar", command=self.cancelar.set, state="disabled")
        self.btn_cancel.pack(side="left", padx=2)
        ttk.Button(bot, text="Abrir carpeta de salida",
                   command=lambda: abrir_carpeta(self.gestor.rutas.salida)).pack(side="left", padx=2)

    # ------------------------------------------------------------ utilidades
    def log(self, msg):
        self.cola.put(("log", msg))

    def _log_ui(self, msg):
        self.txt.configure(state="normal")
        self.txt.insert("end", f"[{datetime.now():%H:%M:%S}] {msg}\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def _procesar_cola(self):
        try:
            while True:
                item = self.cola.get_nowait()
                if item[0] == "log":
                    self._log_ui(item[1])
                elif item[0] == "prog":
                    self.barra.configure(mode="determinate", value=item[1] * 100)
                    self.lbl_estado.configure(text=item[2])
                elif item[0] == "fin":
                    self._fin(*item[1:])
        except queue.Empty:
            pass
        self.after(100, self._procesar_cola)

    def ejecutar(self, trabajo, al_terminar=None, texto="Trabajando..."):
        """Ejecuta 'trabajo' en un hilo; al terminar llama a al_terminar(resultado) en la interfaz."""
        if self.ocupado:
            messagebox.showinfo(APP_NOMBRE, "Hay otra operacion en curso. Espera a que termine.")
            return
        self.ocupado = True
        self.cancelar.clear()
        for b in self.botones:
            b.state(["disabled"])
        self.btn_cancel.configure(state="normal")
        self.barra.configure(mode="indeterminate")
        self.barra.start(12)
        self.lbl_estado.configure(text=texto)

        def run():
            try:
                res = trabajo()
                self.cola.put(("fin", al_terminar, res, None))
            except core.ErrorConsolidador as e:
                self.cola.put(("fin", al_terminar, None, str(e)))
            except Exception as e:  # noqa
                self.cola.put(("fin", al_terminar, None, f"Error inesperado: {type(e).__name__}: {e}"))

        threading.Thread(target=run, daemon=True).start()

    def _fin(self, al_terminar, res, err):
        self.barra.stop()
        self.barra.configure(mode="determinate", value=0)
        self.lbl_estado.configure(text="Listo")
        self.ocupado = False
        for b in self.botones:
            b.state(["!disabled"])
        self.btn_cancel.configure(state="disabled")
        if err:
            self._log_ui(f"ERROR: {err}")
            messagebox.showerror(APP_NOMBRE, err)
        elif al_terminar:
            al_terminar(res) if res is not None else al_terminar()
        self.refrescar_maestros()

    # ----------------------------------------------------------- maestros
    def refrescar_maestros(self, *_):
        sel = self.tree_m.selection()
        self.tree_m.delete(*self.tree_m.get_children())
        for e in self.gestor.estado():
            self.tree_m.insert("", "end", iid=e["clave"], values=(
                e["nombre"], "Cargado" if e["existe"] else "Pendiente",
                f"{e['registros']:,}" if e["registros"] is not None else "-",
                fmt_fecha(e["fecha"]), e["origen"] or "-"))
        if sel and self.tree_m.exists(sel[0]):
            self.tree_m.selection_set(sel[0])

    def _clave_seleccionada(self):
        sel = self.tree_m.selection()
        if not sel:
            messagebox.showinfo(APP_NOMBRE, "Selecciona un maestro de la lista.")
            return None
        return sel[0]

    def actualizar_seleccionado(self):
        c = self._clave_seleccionada()
        if c:
            self.actualizar_maestro(c)

    def actualizar_maestro(self, clave):
        spec = core.MAESTROS[clave]
        tipos = [("CSV", "*.csv"), ("Excel", "*.xlsx *.xls *.xlsm")]
        if spec["zip"]:
            tipos.insert(1, ("ZIP (contiene un CSV)", "*.zip"))
        todos = " ".join(p for _, p in tipos)
        ruta = filedialog.askopenfilename(title=f"Actualizar {spec['nombre']}",
                                          filetypes=[("Archivos compatibles", todos)] + tipos)
        if not ruta:
            return
        self.ejecutar(lambda: self.gestor.actualizar(clave, ruta, self.log),
                      lambda: messagebox.showinfo(APP_NOMBRE, f"{spec['nombre']} actualizado correctamente."),
                      f"Actualizando {spec['nombre']}...")

    def actualizar_excel_tablas(self):
        ruta = filedialog.askopenfilename(title="Excel con tablas maestras (UPS, Establecimiento, Condicion, "
                                                "Profesion)",
                                          filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
        if not ruta:
            return
        self.ejecutar(lambda: self.gestor.importar_excel_tablas(ruta, self.log),
                      lambda: messagebox.showinfo(APP_NOMBRE, "Tablas actualizadas desde el Excel."),
                      "Leyendo Excel...")

    def ver_maestro(self):
        c = self._clave_seleccionada()
        if not c:
            return
        try:
            df = self.gestor.vista_previa(c, 500)
        except core.ErrorConsolidador as e:
            messagebox.showinfo(APP_NOMBRE, str(e))
            return
        win = tk.Toplevel(self)
        win.title(f"{core.MAESTROS[c]['nombre']} - primeras {len(df)} filas")
        win.geometry("1000x500")
        cols = list(df.columns)
        tv = ttk.Treeview(win, columns=cols, show="headings")
        for col in cols:
            tv.heading(col, text=col)
            tv.column(col, width=130, anchor="w")
        for fila in df.itertuples(index=False):
            tv.insert("", "end", values=fila)
        vs = ttk.Scrollbar(win, orient="vertical", command=tv.yview)
        hs = ttk.Scrollbar(win, orient="horizontal", command=tv.xview)
        tv.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        hs.pack(side="bottom", fill="x")
        vs.pack(side="right", fill="y")
        tv.pack(fill="both", expand=True)

    # ------------------------------------------------------------- tramas
    def refrescar_tramas(self, seleccionar_todo=True):
        previas = set(self.tree_t.selection())
        self.tree_t.delete(*self.tree_t.get_children())
        for r in self.gestor.listar_tramas():
            nombre = self.nombres_custom.get(r, core.nombre_por_defecto(r))
            self.tree_t.insert("", "end", iid=r, values=(
                os.path.relpath(r, self.gestor.rutas.tramas), fmt_bytes(os.path.getsize(r)), nombre))
        hijos = self.tree_t.get_children()
        self.tree_t.selection_set(hijos if seleccionar_todo or not previas else [h for h in hijos if h in previas])

    def agregar_tramas(self):
        rutas = filedialog.askopenfilenames(title="Agregar tramas",
                                            filetypes=[("Tramas (CSV o ZIP)", "*.csv *.zip"), ("CSV", "*.csv"),
                                                       ("ZIP", "*.zip")])
        if not rutas:
            return

        def fin(agregadas):
            self.log(f"{len(agregadas)} trama(s) agregada(s).")
            self.refrescar_tramas()
            self.nb.select(1)

        self.ejecutar(lambda: self.gestor.agregar_tramas(rutas, self.log) or [], fin, "Agregando tramas...")

    def quitar_tramas(self):
        sel = self.tree_t.selection()
        if not sel:
            return
        if not messagebox.askyesno(APP_NOMBRE, f"Se eliminaran {len(sel)} trama(s) de la carpeta TRAMAS del "
                                               "programa (el archivo original no se toca). Continuar?"):
            return
        for r in sel:
            try:
                os.remove(r)
            except OSError as e:
                self._log_ui(f"No se pudo eliminar {r}: {e}")
            self.nombres_custom.pop(r, None)
        self.refrescar_tramas(False)

    def editar_nombre(self, ev):
        if self.tree_t.identify_region(ev.x, ev.y) != "cell":
            return
        iid = self.tree_t.identify_row(ev.y)
        if not iid:
            return
        actual = self.nombres_custom.get(iid, core.nombre_por_defecto(iid))
        nuevo = simpledialog.askstring("Nombre del consolidado", "Nombre del archivo de salida:",
                                       initialvalue=actual, parent=self)
        if nuevo and nuevo.strip():
            self.nombres_custom[iid] = nuevo.strip()
            self.tree_t.set(iid, "salida", nuevo.strip())

    def restablecer_nombres(self):
        self.nombres_custom.clear()
        self.refrescar_tramas(False)

    def _modo_cambio(self):
        self.ent_unico.configure(state="normal" if self.modo.get() == "unico" else "disabled")

    def generar(self):
        sel = self.tree_t.selection()
        if not sel:
            messagebox.showinfo(APP_NOMBRE, "Selecciona al menos una trama.")
            return
        items = [(r, self.nombres_custom.get(r, core.nombre_por_defecto(r))) for r in sel]
        modo, unico, bom = self.modo.get(), self.nombre_unico.get(), self.bom.get()

        def trabajo():
            return core.consolidar(
                self.gestor, items, modo, unico, bom, log=self.log,
                progreso=lambda f, t: self.cola.put(("prog", f, t)), cancelar=self.cancelar)

        def fin(generados):
            if generados and messagebox.askyesno(APP_NOMBRE, f"Se generaron {len(generados)} archivo(s) en "
                                                             "SALIDA.\nAbrir la carpeta?"):
                abrir_carpeta(self.gestor.rutas.salida)

        self.ejecutar(trabajo, fin, "Generando consolidados...")
        # El indicador pasa a determinado cuando llega el primer progreso
        self.barra.stop()

    def instrucciones(self):
        messagebox.showinfo("Instrucciones", (
            "1. Datos maestros: revisa que todos esten 'Cargado'. Actualiza cada uno con CSV, ZIP "
            "(Pacientes, Registradores, Personal) o Excel.\n"
            "   Las tablas UPS, Establecimiento, Condicion y Profesion pueden cargarse desde el Excel completo.\n\n"
            "2. Tramas: agrega archivos CSV o ZIP (se copian a la carpeta TRAMAS del programa) y selecciona "
            "las que quieres procesar.\n\n"
            "3. Elige 'separados' (nombre editable por trama) o 'un solo consolidado' y pulsa Generar.\n\n"
            "Los resultados se guardan en la carpeta SALIDA, junto al programa."))


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
