#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ==============================================================================
#   Calorimetro Langavant - Lector de datos SD
#   Conexion TCP/IP o Serie con el M5Stack Core2
#   Requisitos: pip install pyserial
# ==============================================================================

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import socket
import threading
import time
import csv
import os
import json

try:
    import serial
    import serial.tools.list_ports
    SERIAL_OK = True
except ImportError:
    SERIAL_OK = False

COLOR_FONDO    = "#1e1e2e"
COLOR_PANEL    = "#2a2a3e"
COLOR_ACENTO   = "#7aa2f7"
COLOR_OK       = "#9ece6a"
COLOR_ERROR    = "#f7768e"
COLOR_AMARILLO = "#e0af68"
COLOR_TEXTO    = "#cdd6f4"
COLOR_SUBTXT   = "#6c7086"
COLOR_HEADER   = "#313244"

CABECERAS = [
    "Linea","Fecha","Hora",
    "T1(C)","T2(C)","T3(C)","T4(C)","T5(C)","T6(C)","T7(C)",
    "Err1","Err2","Err3","Err4","Err5","Err6","Err7"
]

# ==============================================================================
#   PARSEO DE TRAMAS
# ==============================================================================
def parsear_EST(t):
    try:
        t = t.strip().replace("\r","").replace("\n","")
        p = t.split(";")
        if p[0] != "EST" or len(p) < 13:
            return None
        return {
            "estado"   : int(p[1]),
            "intervalo": int(p[2]),
            "filas"    : int(p[3]),
            "esclavos" : int(p[4]),
            "cargando" : bool(int(p[5])),
            "bateria"  : int(p[6]),
            "hora"     : f"{p[7]}:{p[8]}:{p[9]}",
            "fecha"    : f"{p[10]}/{p[11]}/{p[12]}",
        }
    except Exception as e:
        print(f"Error parsear_EST: {e}  |  trama: [{t}]")
        return None

def parsear_LFI(t, n):
    try:
        t = t.strip().replace("\r","").replace("\n","")
        p = t.split(";")
        if p[0] != "LFI" or len(p) < 21:
            return None
        fecha = f"{p[1]}/{p[2]}/{p[3]}"
        hora  = f"{p[4]}:{p[5]}:{p[6]}"
        return [str(n), fecha, hora] + p[7:]
    except:
        return None
def parsear_LMAC(t):
    try:
        t = t.strip().replace("\r","").replace("\n","")
        p = t.split(";")
        if len(p) < 2: return None
        # Acepta tanto "LMAC" como cualquier cabecera
        return [m.strip() for m in p[1:] if m.strip()]
    except:
        return None

def parsear_LSO(t):
    try:
        t = t.strip().replace("\r","").replace("\n","")
        p = t.split(";")
        if p[0] != "LSO" or len(p) < 8:
            return None
        return [p[1],p[2],p[3],p[4],p[5],p[6],p[7]]
    except:
        return None

def esclavos_lista(m):
    return [i+1 for i in range(8) if m & (1 << i)]

# ==============================================================================
#   CONFIG
# ==============================================================================
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

def guardar_config(datos):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(datos, f)
    except Exception as e:
        print(f"Error guardando config: {e}")

def cargar_config():
    try:
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    except:
        return {}

# ==============================================================================
#   CONEXION
# ==============================================================================
class Conexion:
    def __init__(self):
        self.sock = None
        self.ser  = None
        self.modo = None
        self.lock = threading.Lock()

    def conectar_tcp(self, ip, pto):
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.settimeout(3)
        self.sock.connect((ip, int(pto)))
        self.modo = "tcp"

    def conectar_serie(self, pto, baud):
        self.ser = serial.Serial(pto, int(baud), timeout=3)
        self.modo = "serie"

    def desconectar(self):
        try:
            if self.sock: self.sock.close()
            if self.ser:  self.ser.close()
        except:
            pass
        self.sock = self.ser = self.modo = None

    def enviar(self, cmd):
        d = (cmd + "\r\n").encode("ascii")
        if self.modo == "tcp":
            self.sock.sendall(d)
        else:
            self.ser.write(d)

    def recibir_linea(self):
        if self.modo == "tcp":
            buf = b""
            while True:
                c = self.sock.recv(1)
                if not c: break
                buf += c
                if c == b"\n": break
            return buf.decode("ascii", errors="ignore").strip()
        return self.ser.readline().decode("ascii", errors="ignore").strip()

    def comando(self, cmd):
        with self.lock:
            self.enviar(cmd)
            resp = self.recibir_linea()
            if self.modo == "serie":
                time.sleep(0.002)
            return resp

    @property
    def conectado(self):
        return self.modo is not None

# ==============================================================================
#   APLICACION
# ==============================================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Calorimetro Langavant  -  Lector SD  v1.0")
        self.state("zoomed")
        self.configure(bg=COLOR_FONDO)
        self.conn        = Conexion()
        self.filas_sd    = 0
        self.datos       = []
        self.descargando = False
        self.leyendo_continuo = False
        self._ui()

        cfg = cargar_config()
        if "ip"     in cfg: self.ip_var.set(cfg["ip"])
        if "puerto" in cfg: self.pto_var.set(cfg["puerto"])

        self.ip_var.trace_add("write",
            lambda *a: guardar_config({"ip": self.ip_var.get(),
                                       "puerto": self.pto_var.get()}))
        self.pto_var.trace_add("write",
            lambda *a: guardar_config({"ip": self.ip_var.get(),
                                       "puerto": self.pto_var.get()}))
        self._actualizar_puertos()

    # --------------------------------------------------------------------------
    #   HELPERS
    # --------------------------------------------------------------------------
    def _lbl(self, p, t):
        return tk.Label(p, text=t, bg=p.cget("bg"), fg="white",
                        font=("Consolas", 9))

    def _entry(self, p, v, w):
        return tk.Entry(p, textvariable=v, width=w, bg="#313244",
                        fg=COLOR_TEXTO, insertbackground=COLOR_TEXTO,
                        relief="flat", font=("Consolas", 10))

    def _log(self, m):
        self.log_var.set(m)
        import datetime
        hora = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{hora}]  {m}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _borrar_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _modo_cambio(self):
        if self.modo_var.get() == "tcp":
            self.frame_tcp.grid()
            self.frame_serie.grid_remove()
        else:
            self.frame_tcp.grid_remove()
            self.frame_serie.grid()

    def _actualizar_puertos(self):
        if SERIAL_OK:
            pts = [p.device for p in serial.tools.list_ports.comports()]
            self.combo_com["values"] = pts
            if pts and not self.com_var.get():
                self.com_var.set(pts[0])
        self.after(3000, self._actualizar_puertos)

    # --------------------------------------------------------------------------
    #   UI PRINCIPAL
    # --------------------------------------------------------------------------
    def _ui(self):
        tk.Label(self, text="CALORIMETRO LANGAVANT",
                 bg=COLOR_FONDO, fg=COLOR_ACENTO,
                 font=("Consolas", 15, "bold")).pack(pady=(12,0))
        sf = tk.Frame(self, bg=COLOR_FONDO)
        sf.pack(pady=(0,8))
        tk.Label(sf, text="M5Stack Core2  |  ESP-NOW Master  |  Ethernet TCP/IP  |  ",
                 bg=COLOR_FONDO, fg=COLOR_SUBTXT,
                 font=("Consolas", 9)).pack(side="left")
        self.version_var = tk.StringVar(value="v?.?")
        tk.Label(sf, textvariable=self.version_var,
                 bg=COLOR_FONDO, fg=COLOR_ACENTO,
                 font=("Consolas", 9, "bold")).pack(side="left")
        tk.Label(sf, text="     UUID: ",
                 bg=COLOR_FONDO, fg=COLOR_SUBTXT,
                 font=("Consolas", 9)).pack(side="left")
        self.uuid_var = tk.StringVar(value="--")
        tk.Label(sf, textvariable=self.uuid_var,
                 bg=COLOR_FONDO, fg=COLOR_ACENTO,
                 font=("Consolas", 9, "bold")).pack(side="left")

        top = tk.Frame(self, bg=COLOR_FONDO)
        top.pack(fill="x", padx=12, pady=4)
        self._panel_conexion(top)
        self._panel_estado(top)
        self._panel_tabla()
        self._panel_sondas()

        f = tk.Frame(self, bg=COLOR_HEADER)
        f.pack(fill="x", side="bottom")
        tk.Label(f, text=" Log:", bg=COLOR_HEADER, fg=COLOR_SUBTXT,
                 font=("Consolas", 8, "bold")).pack(side="left", padx=6, pady=2)
        tk.Button(f, text="🗑", command=self._borrar_log,
                  bg=COLOR_HEADER, fg=COLOR_SUBTXT,
                  font=("Consolas", 9), relief="flat",
                  cursor="hand2").pack(side="left", padx=(0,4))
        self.log_text = tk.Text(f, height=4, bg=COLOR_HEADER, fg=COLOR_TEXTO,
                                font=("Consolas", 8), relief="flat",
                                state="disabled", wrap="word")
        self.log_text.pack(side="left", fill="x", expand=True, padx=4, pady=2)
        sb = ttk.Scrollbar(f, orient="vertical", command=self.log_text.yview)
        sb.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=sb.set)
        self.log_var = tk.StringVar(value="Listo.")

    # --------------------------------------------------------------------------
    #   PANEL CONEXION
    # --------------------------------------------------------------------------
    def _panel_conexion(self, padre):
        f = tk.LabelFrame(padre, text="  Conexion  ", bg=COLOR_PANEL,
                          fg=COLOR_ACENTO, font=("Consolas",9,"bold"),
                          relief="flat")
        f.pack(side="left", fill="y", padx=(0,8), ipadx=10, ipady=8)

        self.modo_var = tk.StringVar(value="tcp")
        mf = tk.Frame(f, bg=COLOR_PANEL)
        mf.grid(row=0, column=0, columnspan=4, sticky="w", pady=4)
        for txt, val in [("TCP/IP","tcp"),("Serie (COM)","serie")]:
            tk.Radiobutton(mf, text=txt, variable=self.modo_var, value=val,
                           bg=COLOR_PANEL, fg=COLOR_TEXTO,
                           selectcolor=COLOR_FONDO,
                           activebackground=COLOR_PANEL,
                           command=self._modo_cambio).pack(side="left", padx=4)

        self.frame_tcp = tk.Frame(f, bg=COLOR_PANEL)
        self.frame_tcp.grid(row=1, column=0, columnspan=4, sticky="w", pady=2)
        self._lbl(self.frame_tcp,"IP:").grid(row=0,column=0,sticky="w")
        self.ip_var = tk.StringVar(value="192.168.0.250")
        self._entry(self.frame_tcp, self.ip_var, 16).grid(row=0,column=1,padx=4)
        self._lbl(self.frame_tcp,"Puerto:").grid(row=0,column=2,sticky="w")
        self.pto_var = tk.StringVar(value="20256")
        self._entry(self.frame_tcp, self.pto_var, 7).grid(row=0,column=3,padx=4)

        # Frame cambio IP (visible siempre, fuera de frame_tcp y frame_serie)
        fip = tk.Frame(f, bg=COLOR_PANEL)
        fip.grid(row=3, column=0, columnspan=4, sticky="w", pady=(4,0), padx=4)
        self._lbl(fip, "Nueva IP:").grid(row=0,column=0,sticky="w")
        self.nueva_ip_var = tk.StringVar(value="")
        self._entry(fip, self.nueva_ip_var, 16).grid(row=0,column=1,padx=4)
        tk.Button(fip, text="  Cambiar IP",
                  command=self._enviar_nueva_ip,
                  bg=COLOR_OK, fg="#1e1e2e",
                  font=("Consolas",9,"bold"),
                  relief="flat", padx=8,
                  cursor="hand2").grid(row=0,column=2,padx=4,sticky="w")

        self.frame_serie = tk.Frame(f, bg=COLOR_PANEL)
        self.frame_serie.grid(row=1,column=0,columnspan=4,sticky="w",pady=2)
        self.frame_serie.grid_remove()
        self._lbl(self.frame_serie,"Puerto:").grid(row=0,column=0,sticky="w")
        self.com_var = tk.StringVar()
        self.combo_com = ttk.Combobox(self.frame_serie, textvariable=self.com_var,
                                      width=10, state="readonly")
        self.combo_com.grid(row=0, column=1, padx=4)
        self._lbl(self.frame_serie,"Baudios:").grid(row=0,column=2,sticky="w")
        self.baud_var = tk.StringVar(value="115200")
        ttk.Combobox(self.frame_serie, textvariable=self.baud_var, width=9,
                     values=["9600","19200","57600","115200",
                             "230400","460800","921600"],
                     state="readonly").grid(row=0,column=3,padx=4)
        tk.Button(self.frame_serie, text=" \u21ba ",
                  command=self._actualizar_puertos,
                  bg=COLOR_PANEL, fg=COLOR_ACENTO,
                  font=("Consolas",9), relief="flat",
                  cursor="hand2").grid(row=0, column=4, padx=4)

        bf = tk.Frame(f, bg=COLOR_PANEL)
        bf.grid(row=2, column=0, columnspan=4, pady=(10,2), sticky="w")
        self.btn_con = tk.Button(bf, text="  Conectar",
                                 command=self._conectar, bg=COLOR_ACENTO,
                                 fg="#1e1e2e", font=("Consolas",9,"bold"),
                                 relief="flat", padx=10, cursor="hand2")
        self.btn_con.pack(side="left")
        self.btn_des = tk.Button(bf, text="  Desconectar",
                                 command=self._desconectar, bg=COLOR_ERROR,
                                 fg="#1e1e2e", font=("Consolas",9,"bold"),
                                 relief="flat", padx=10, cursor="hand2",
                                 state="disabled")
        self.btn_des.pack(side="left", padx=(8,0))
        self.led_var = tk.StringVar(value="  DESCONECTADO")
        self.led_lbl = tk.Label(bf, textvariable=self.led_var,
                                bg=COLOR_PANEL, fg=COLOR_ERROR,
                                font=("Consolas",9,"bold"))
        self.led_lbl.pack(side="left", padx=(16,0))

    # --------------------------------------------------------------------------
    #   PANEL ESTADO (?EST) + ESTA
    # --------------------------------------------------------------------------
    def _panel_estado(self, padre):
        f = tk.LabelFrame(padre, text="  Estado del equipo  ",
                          bg=COLOR_PANEL, fg=COLOR_ACENTO,
                          font=("Consolas",9,"bold"), relief="flat")
        f.pack(side="left", fill="both", expand=True, ipadx=10, ipady=8)

        # ── Columna izquierda: ?EST ───────────────────────────────────────────
        left = tk.Frame(f, bg=COLOR_PANEL)
        left.pack(side="left", fill="y", padx=(0,10))

        tk.Label(left, text="?EST", bg=COLOR_PANEL, fg=COLOR_ACENTO,
                 font=("Consolas",9,"bold")).grid(row=0, column=0,
                 columnspan=4, sticky="w", padx=(8,0), pady=(0,4))

        self.est = {}
        for i,(lbl,key) in enumerate([
                ("Estado:","estado"), ("Fecha/Hora:","fechahora"),
                ("Filas SD:","filas"), ("Intervalo:","intervalo"),
                ("Bateria:","bateria"), ("Esclavos:","esclavos")]):
            r, c = divmod(i, 2)
            self._lbl(left,lbl).grid(row=r+1,column=c*2,sticky="w",
                                     padx=(8,2),pady=3)
            v = tk.StringVar(value="---")
            self.est[key] = v
            tk.Label(left, textvariable=v, bg=COLOR_PANEL, fg=COLOR_AMARILLO,
                     font=("Consolas",10,"bold")).grid(
                         row=r+1,column=c*2+1,sticky="w",padx=(0,20))

        bf = tk.Frame(left, bg=COLOR_PANEL)
        bf.grid(row=4, column=0, columnspan=4, sticky="w", pady=(8,0), padx=8)
        tk.Button(bf, text="  Leer estado",
                  command=self._leer_estado,
                  bg=COLOR_OK, fg="#1e1e2e",
                  font=("Consolas",9,"bold"),
                  relief="flat", padx=10,
                  cursor="hand2").pack(side="left")

        bf3 = tk.Frame(left, bg=COLOR_PANEL)
        bf3.grid(row=5, column=0, columnspan=4, sticky="w", pady=(6,0), padx=8)
        self._lbl(bf3, "Idioma:").pack(side="left", padx=(0,6))
        self.idioma_var = tk.StringVar(value="1 - Español")
        ttk.Combobox(bf3, textvariable=self.idioma_var, width=12,
                     values=["0 - Ingles", "1 - Español"],
                     state="readonly").pack(side="left")
        tk.Button(bf3, text="  Enviar",
                  command=self._enviar_idioma,
                  bg=COLOR_OK, fg="#1e1e2e",
                  font=("Consolas",9,"bold"),
                  relief="flat", padx=10,
                  cursor="hand2").pack(side="left", padx=(6,0))
        tk.Button(bf, text="  Leer linea",
                  command=self._leer_linea_manual,
                  bg=COLOR_OK, fg="#1e1e2e",
                  font=("Consolas",9,"bold"),
                  relief="flat", padx=10,
                  cursor="hand2").pack(side="left", padx=(20,6))
        self.linea_var = tk.StringVar(value="1")
        self._entry(bf, self.linea_var, 6).pack(side="left")
        tk.Button(bf, text="  Sincronizar hora",
                  command=self._sincronizar_hora,
                  bg=COLOR_OK, fg="#1e1e2e",
                  font=("Consolas",9,"bold"),
                  relief="flat", padx=10,
                  cursor="hand2").pack(side="left", padx=(6,0))

        # Separador vertical
        tk.Frame(f, bg=COLOR_ACENTO, width=1).pack(side="left",
                 fill="y", padx=(0,10), pady=4)

        # ── Columna centro: ESTA ──────────────────────────────────────────────
        right = tk.Frame(f, bg=COLOR_PANEL)
        right.pack(side="left", fill="y")

        tk.Label(right, text="ESTA", bg=COLOR_PANEL, fg=COLOR_ACENTO,
                 font=("Consolas",9,"bold")).grid(row=0, column=0,
                 columnspan=4, sticky="w", padx=(8,0), pady=(0,4))

        self._lbl(right,"Modo:").grid(row=1,column=0,sticky="w",padx=(8,2),pady=3)
        self.esta_modo_var = tk.StringVar(value="0 - Parar")
        ttk.Combobox(right, textvariable=self.esta_modo_var, width=14,
                     values=["0 - Parar","1 - Calibrando","2 - Ensayar"],
                     state="readonly").grid(row=1,column=1,sticky="w",padx=4)

        self._lbl(right,"Intervalo SD (seg):").grid(row=2,column=0,
                  sticky="w",padx=(8,2),pady=3)
        self.esta_intervalo_var = tk.StringVar(value="60")
        self._entry(right, self.esta_intervalo_var, 8).grid(row=2,column=1,
                    sticky="w",padx=4)

        self._lbl(right,"Sondas activas:").grid(row=3,column=0,
                  sticky="w",padx=(8,2),pady=3)
        sc = tk.Frame(right, bg=COLOR_PANEL)
        sc.grid(row=3, column=1, sticky="w", padx=4)
        self.sonda_checks = []
        col_izq = tk.Frame(sc, bg=COLOR_PANEL)
        col_izq.pack(side="left", anchor="n", padx=(0,8))
        col_der = tk.Frame(sc, bg=COLOR_PANEL)
        col_der.pack(side="left", anchor="n")
        for i, nombre in enumerate(["Ref","S1","S2","S3","S4","S5","S6"]):
            v = tk.BooleanVar(value=True)
            self.sonda_checks.append(v)
            parent = col_izq if i < 4 else col_der
            tk.Checkbutton(parent, text=nombre, variable=v,
                           bg=COLOR_PANEL, fg="white",
                           selectcolor=COLOR_FONDO,
                           activebackground=COLOR_PANEL,
                           font=("Consolas",8),
                           command=self._actualizar_estado_entries_mac
                           ).pack(anchor="w", padx=1)
        bf2 = tk.Frame(right, bg=COLOR_PANEL)
        bf2.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8,0), padx=8)
        tk.Button(bf2, text="  Enviar estado",
                  command=self._enviar_estado,
                  bg=COLOR_OK, fg="#1e1e2e",
                  font=("Consolas",9,"bold"),
                  relief="flat", padx=10,
                  cursor="hand2").pack(side="left")
        # Separador vertical
        tk.Frame(f, bg=COLOR_ACENTO, width=1).pack(side="left",
                 fill="y", padx=(10,10), pady=4)

        # ── Columna derecha: MACs ─────────────────────────────────────────────
        mac_frame = tk.Frame(f, bg=COLOR_PANEL)
        mac_frame.pack(side="left", fill="y")

        tk.Label(mac_frame, text="Leer MACs (?LMAC)", bg=COLOR_PANEL, fg=COLOR_ACENTO,
                 font=("Consolas",9,"bold")).pack(anchor="w", padx=(8,0),
                 pady=(0,4))

        self.mac_vars = []
        self.mac_entries = []
        for i in range(7):
            fila = tk.Frame(mac_frame, bg=COLOR_PANEL)
            fila.pack(anchor="w", pady=1)
            nombre = "Ref" if i == 0 else f"S{i} "
            tk.Label(fila, text=f"{nombre}:", bg=COLOR_PANEL,
                     fg=COLOR_ACENTO, font=("Consolas",9,"bold"),
                     width=4, anchor="w").pack(side="left", padx=(8,4))
            v = tk.StringVar(value="")
            self.mac_vars.append(v)
            e = tk.Entry(fila, textvariable=v, width=14,
                         bg="#313244", fg=COLOR_AMARILLO,
                         insertbackground=COLOR_AMARILLO,
                         relief="flat", font=("Consolas",9))
            e.pack(side="left")
            self.mac_entries.append(e)

        tk.Button(mac_frame, text="  Leer MACs",
                  command=self._leer_macs,
                  bg=COLOR_OK, fg="#1e1e2e",
                  font=("Consolas",9,"bold"),
                  relief="flat", padx=10,
                  cursor="hand2").pack(anchor="w", padx=8, pady=(8,0))
        tk.Button(mac_frame, text="  Enviar MACs",
                  command=self._enviar_macs,
                  bg=COLOR_OK, fg="#1e1e2e",
                  font=("Consolas",9,"bold"),
                  relief="flat", padx=10,
                  cursor="hand2").pack(anchor="w", padx=8, pady=(4,0))
        
    # --------------------------------------------------------------------------
    #   PANEL TABLA
    # --------------------------------------------------------------------------
    def _panel_tabla(self):
        mid = tk.Frame(self, bg=COLOR_FONDO)
        mid.pack(fill="both", expand=True, padx=12, pady=4)

        bdf = tk.Frame(mid, bg=COLOR_PANEL)
        bdf.pack(fill="x", pady=(0,6), ipadx=8, ipady=6)
        tk.Button(bdf, text="  Descargar TODO",
                  command=self._descargar_todo, bg=COLOR_ACENTO,
                  fg="#1e1e2e", font=("Consolas",10,"bold"),
                  relief="flat", padx=14, cursor="hand2").pack(side="left",padx=8)
        self._lbl(bdf,"  Desde:").pack(side="left",padx=(16,4))
        self.desde_var = tk.StringVar(value="1")
        self._entry(bdf, self.desde_var, 6).pack(side="left")
        self._lbl(bdf,"  Hasta:").pack(side="left",padx=(8,4))
        self.hasta_var = tk.StringVar(value="")
        self._entry(bdf, self.hasta_var, 6).pack(side="left")
        tk.Button(bdf, text="  Descargar rango",
                  command=self._descargar_rango, bg=COLOR_PANEL,
                  fg=COLOR_ACENTO, font=("Consolas",9), relief="flat",
                  padx=10, cursor="hand2",
                  highlightbackground=COLOR_ACENTO,
                  highlightthickness=1).pack(side="left",padx=8)
        self.btn_parar = tk.Button(bdf, text="  Parar",
                                   command=self._parar_descarga,
                                   bg=COLOR_ERROR, fg="#1e1e2e",
                                   font=("Consolas",9,"bold"), relief="flat",
                                   padx=10, cursor="hand2", state="disabled")
        self.btn_parar.pack(side="left", padx=4)
        tk.Button(bdf, text="  Ver grafica", command=self._ver_grafica,
                  bg="#414868", fg=COLOR_TEXTO, font=("Consolas",9),
                  relief="flat", padx=10,
                  cursor="hand2").pack(side="right",padx=8)
        tk.Button(bdf, text="  Exportar CSV", command=self._exportar_csv,
                  bg="#414868", fg=COLOR_TEXTO, font=("Consolas",9),
                  relief="flat", padx=10,
                  cursor="hand2").pack(side="right",padx=8)
        tk.Button(bdf, text="  Limpiar", command=self._limpiar_tabla,
                  bg="#414868", fg=COLOR_TEXTO, font=("Consolas",9),
                  relief="flat", padx=10,
                  cursor="hand2").pack(side="right",padx=4)

        pf = tk.Frame(mid, bg=COLOR_FONDO)
        pf.pack(fill="x", pady=(0,4))
        self.prog_var = tk.DoubleVar(value=0)
        ttk.Progressbar(pf, variable=self.prog_var,
                        maximum=100, length=500).pack(side="left",padx=4)
        self.prog_lbl = tk.Label(pf, text="", bg=COLOR_FONDO,
                                 fg=COLOR_TEXTO, font=("Consolas",9))
        self.prog_lbl.pack(side="left", padx=8)

        tf = tk.Frame(mid, bg=COLOR_FONDO)
        tf.pack(fill="both", expand=True)
        style = ttk.Style()
        style.theme_use("default")
        style.configure("C.Treeview", background=COLOR_PANEL,
                        foreground=COLOR_TEXTO, fieldbackground=COLOR_PANEL,
                        rowheight=24, font=("Consolas",9))
        style.configure("C.Treeview.Heading", background=COLOR_HEADER,
                        foreground=COLOR_ACENTO, font=("Consolas",9,"bold"),
                        relief="flat")
        style.map("C.Treeview",
                  background=[("selected","#364a82")],
                  foreground=[("selected","white")])

        self.tree = ttk.Treeview(tf, columns=CABECERAS,
                                 show="headings", style="C.Treeview")
        anchos = {"Linea":48,"Fecha":72,"Hora":68,
                  "T1(C)":78,"T2(C)":78,"T3(C)":78,"T4(C)":78,
                  "T5(C)":78,"T6(C)":78,"T7(C)":78,
                  "Err1":44,"Err2":44,"Err3":44,"Err4":44,
                  "Err5":44,"Err6":44,"Err7":44}
        for col in CABECERAS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=anchos.get(col,60),
                             anchor="center", stretch=False)
        sb_y = ttk.Scrollbar(tf, orient="vertical",   command=self.tree.yview)
        sb_x = ttk.Scrollbar(tf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right", fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)
        self.tree.tag_configure("par",   background="#252535")
        self.tree.tag_configure("impar", background=COLOR_PANEL)
        self.tree.tag_configure("err",   background="#3d1f2a",
                                         foreground=COLOR_ERROR)

    # --------------------------------------------------------------------------
    #   PANEL SONDAS
    # --------------------------------------------------------------------------
    def _panel_sondas(self):
        f = tk.LabelFrame(self, text="  Lecturas de Sondas  (?LSO)  ",
                          bg=COLOR_PANEL, fg=COLOR_ACENTO,
                          font=("Consolas",9,"bold"), relief="flat")
        f.pack(fill="x", padx=12, pady=(0,4), ipadx=8, ipady=6)

        self.sonda_vars = []
        sf = tk.Frame(f, bg=COLOR_PANEL)
        sf.pack(side="left", fill="x", expand=True)
        for i in range(7):
            col = tk.Frame(sf, bg=COLOR_PANEL)
            col.pack(side="left", padx=8)
            lbl = "Referencia" if i == 0 else f"Sonda {i}"
            tk.Label(col, text=lbl, bg=COLOR_PANEL, fg=COLOR_ACENTO,
                     font=("Consolas",9,"bold")).pack()
            v = tk.StringVar(value="---")
            self.sonda_vars.append(v)
            df = tk.Frame(col, bg="#18181b", relief="flat", bd=1)
            df.pack(pady=2)
            tk.Label(df, textvariable=v, bg="#18181b", fg="white",
                     font=("Consolas",14,"bold"),
                     anchor="center").pack(side="left", padx=(6,0))
            tk.Label(df, text="\u00b0C", bg="#18181b", fg="white",
                     font=("Consolas",14,"bold")).pack(side="left", padx=(2,6))

        bf = tk.Frame(f, bg=COLOR_PANEL)
        bf.pack(side="left", padx=(20,0))
        tk.Button(bf, text="  Leer sondas",
                  command=self._leer_sondas,
                  bg=COLOR_OK, fg="white",
                  font=("Consolas",9,"bold"),
                  relief="flat", padx=10,
                  cursor="hand2").pack(pady=(0,6))

        self.lec_continua_var = tk.BooleanVar(value=False)
        tk.Checkbutton(bf, text="Lectura continua",
                       variable=self.lec_continua_var,
                       command=self._toggle_continua,
                       bg=COLOR_PANEL, fg="white",
                       selectcolor=COLOR_FONDO,
                       activebackground=COLOR_PANEL,
                       font=("Consolas",9)).pack()

        tk.Label(bf, text="Intervalo (seg):", bg=COLOR_PANEL,
                 fg="white", font=("Consolas",9)).pack(pady=(6,2))
        self.intervalo_lec_var = tk.StringVar(value="2")
        self._entry(bf, self.intervalo_lec_var, 5).pack()

    # --------------------------------------------------------------------------
    #   CONEXION
    # --------------------------------------------------------------------------
    def _conectar(self):
        try:
            if self.modo_var.get() == "tcp":
                self.conn.conectar_tcp(self.ip_var.get(), self.pto_var.get())
                self._log(f"Conectado TCP  {self.ip_var.get()}:{self.pto_var.get()}")
            else:
                if not SERIAL_OK:
                    messagebox.showerror("Error","pip install pyserial"); return
                self.conn.conectar_serie(self.com_var.get(), self.baud_var.get())
                self._log(f"Conectado Serie {self.com_var.get()}")
            self.led_var.set("  CONECTADO")
            self.led_lbl.config(fg=COLOR_OK)
            self.btn_con.config(state="disabled")
            self.btn_des.config(state="normal")
            self._leer_estado()
            self._leer_version()
            self._leer_uuid()
            self._leer_macs()
            self.after(500, self._actualizar_estado_entries_mac)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._log(f"Error: {e}")

    def _desconectar(self):
        self._parar_descarga()
        self.conn.desconectar()
        self.led_var.set("  DESCONECTADO")
        self.led_lbl.config(fg=COLOR_ERROR)
        self.btn_con.config(state="normal")
        self.btn_des.config(state="disabled")
        self._log("Desconectado.")

    def _reconectar(self):
        if not self.conn.conectado:
            return
        self._log("Conexion perdida, reconectando...")
        try:
            self.conn.desconectar()
            time.sleep(1)
            if self.modo_var.get() == "tcp":
                self.conn.conectar_tcp(self.ip_var.get(), self.pto_var.get())
            else:
                self.conn.conectar_serie(self.com_var.get(), self.baud_var.get())
            self._log("Reconectado OK")
            self.led_var.set("  CONECTADO")
            self.led_lbl.config(fg=COLOR_OK)
            self._leer_estado()
            self._leer_version()
            self._leer_uuid()
            self._leer_macs()
            self.after(500, self._actualizar_estado_entries_mac)
        except Exception as e:
            self._log(f"Error reconectando: {e}")
            self.led_var.set("  DESCONECTADO")
            self.led_lbl.config(fg=COLOR_ERROR)
            self.btn_con.config(state="normal")
            self.btn_des.config(state="disabled")
        

    # --------------------------------------------------------------------------
    #   COMANDOS
    # --------------------------------------------------------------------------
    def _leer_estado(self):
        if not self.conn.conectado:
            messagebox.showwarning("Sin conexion","Conecta primero."); return
        def hilo():
            try:
                resp = self.conn.comando("?EST")
                self.after(0, self._log, f"<- {resp}")
                d = parsear_EST(resp)
                if not d:
                    self.after(0, self._log, f"Respuesta inesperada: {resp}")
                    return
                self.filas_sd = d["filas"]
                et = {0:"DETENIDO",1:"CALIBRANDO",2:"ENSAYANDO"}
                self.after(0, self.est["estado"].set, et.get(d["estado"],"?"))
                self.after(0, self.est["fechahora"].set,
                           f"{d['fecha']}  {d['hora']}")
                self.after(0, self.est["filas"].set,
                           f"{d['filas']:,} lineas")
                self.after(0, self.est["intervalo"].set,
                           f"{d['intervalo']} seg")
                self.after(0, self.est["bateria"].set,
                           f"{d['bateria']} %")
                self.after(0, self.est["esclavos"].set,
                           str(esclavos_lista(d["esclavos"])))
                self.after(0, self.hasta_var.set, str(self.filas_sd))
                self.after(0, self.est["esclavos"].set,
                           str(esclavos_lista(d["esclavos"])))
                self.after(0, self.hasta_var.set, str(self.filas_sd))

                # NUEVO - Actualizar controles ESTA con los valores del equipo
                modos = {0:"0 - Parar", 1:"1 - Calibrando", 2:"2 - Ensayar"}
                self.after(0, self.esta_modo_var.set,
                           modos.get(d["estado"], "0 - Parar"))
                self.after(0, self.esta_intervalo_var.set,
                           str(d["intervalo"]))
                mascara = d["esclavos"]
                for i, v in enumerate(self.sonda_checks):
                    self.after(0, v.set, bool(mascara & (1 << i)))
                    
            except Exception as e:
                self.after(0, self._log, f"Error leyendo estado: {e}")
                self.after(0, self._reconectar)
        threading.Thread(target=hilo, daemon=True).start()
    def _leer_version(self):
        def hilo():
            try:
                resp = self.conn.comando("?V")
                self.after(0, self._log, f"<- {resp}")
                self.after(0, self.version_var.set, resp.strip().replace(";", " "))
            except Exception as e:
                self.after(0, self._log, f"Error leyendo version: {e}")
                self.after(0, self._reconectar)
        threading.Thread(target=hilo, daemon=True).start()

    def _leer_uuid(self):
        def hilo():
            try:
                resp = self.conn.comando("?UUID")
                self.after(0, self._log, f"<- {resp}")
                t = resp.strip().replace("\r","").replace("\n","")
                p = t.split(";")
                if len(p) >= 2:
                    self.after(0, self.uuid_var.set, p[1].strip())
            except Exception as e:
                self.after(0, self._log, f"Error leyendo UUID: {e}")
                self.after(0, self._reconectar)
        threading.Thread(target=hilo, daemon=True).start()

    def _leer_macs(self):
        if not self.conn.conectado:
            messagebox.showwarning("Sin conexion","Conecta primero."); return
        def hilo():
            try:
                resp = self.conn.comando("?LMAC")
                #self.after(0, self._log, f"[resp}")
                self.after(0, self._log, f"<- {resp}")
                macs = parsear_LMAC(resp)
                if macs:
                    self.after(0, self._actualizar_macs, macs)
                else:
                    self.after(0, self._log, f"Respuesta inesperada LMAC: [{resp}]")
            except Exception as e:
                self.after(0, self._log, f"Error leyendo MACs: {e}")
                self.after(0, self._reconectar)
        threading.Thread(target=hilo, daemon=True).start()

    def _actualizar_macs(self, macs):
        for v in self.mac_vars:
            v.set("")
        for i, mac in enumerate(macs):
            if i >= len(self.mac_vars): break
            self.mac_vars[i].set(mac.strip())
        self._actualizar_estado_entries_mac()

    def _actualizar_estado_entries_mac(self):
        for i, e in enumerate(self.mac_entries):
            if self.sonda_checks[i].get():
                e.config(state="normal")
            else:
                e.config(state="disabled")
            
    def _enviar_macs(self):
        if not self.conn.conectado:
            messagebox.showwarning("Sin conexion","Conecta primero."); return
        def hilo():
            try:
                cmd = "EMAC"
                enviadas = 0
                for i, v in enumerate(self.mac_vars):
                    # Solo sondas habilitadas
                    if not self.sonda_checks[i].get():
                        continue
                    mac = v.get().strip().replace(":","").replace("-","")
                    if len(mac) != 12:
                        self.after(0, self._log,
                                   f"MAC {i} invalida: [{mac}]")
                        return
                    cmd += f"#{mac}"
                    enviadas += 1
                if enviadas == 0:
                    self.after(0, messagebox.showwarning,
                               "Sin sondas", "No hay sondas habilitadas.")
                    return
                cmd += "#"
                self.after(0, self._log, f"-> {cmd}")
                resp = self.conn.comando(cmd)
                self.after(0, self._log, f"<- {resp}")
                if "OK" in resp:
                    self.after(0, messagebox.showinfo,
                               "MACs enviadas",
                               f"{enviadas} MACs enviadas correctamente.")
                else:
                    self.after(0, messagebox.showerror,
                               "Error", f"Respuesta inesperada:\n{resp}")
            except Exception as e:
                self.after(0, self._log, f"Error enviando MACs: {e}")
        threading.Thread(target=hilo, daemon=True).start()
            
    def _leer_linea_manual(self):
        if not self.conn.conectado:
            messagebox.showwarning("Sin conexion","Conecta primero."); return
        try:
            n = int(self.linea_var.get())
            resp = self.conn.comando(f"?LFI#{n}#")
            self._log(f"<- {resp}")
            fila = parsear_LFI(resp, n)
            if fila: self._insertar_fila(fila)
            else: self._log(f"Error linea {n}: {resp}")
        except Exception as e:
            self._log(f"Error: {e}")

    def _sincronizar_hora(self):
        if not self.conn.conectado:
            messagebox.showwarning("Sin conexion","Conecta primero."); return
        def hilo():
            try:
                import datetime
                ahora = datetime.datetime.now()
                cmd = (f"TIME#{ahora.hour:02d}#"
                       f"{ahora.minute:02d}#"
                       f"{ahora.second:02d}#"
                       f"{ahora.day:02d}#"
                       f"{ahora.month:02d}#"
                       f"{ahora.year % 100:02d}#")
                self.after(0, self._log, f"-> {cmd}")
                resp = self.conn.comando(cmd)
                self.after(0, self._log, f"<- {resp}")
                if "OK" in resp:
                    self.after(0, messagebox.showinfo,
                               "Hora sincronizada",
                               f"Hora del PC enviada correctamente:\n{cmd}")
            except Exception as e:
                self.after(0, self._log, f"Error sincronizando hora: {e}")
        threading.Thread(target=hilo, daemon=True).start()

    def _enviar_estado(self):
        if not self.conn.conectado:
            messagebox.showwarning("Sin conexion","Conecta primero."); return
        def hilo():
            try:
                modo = int(self.esta_modo_var.get().split(" ")[0])
                intervalo = int(self.esta_intervalo_var.get())
                mascara = 0
                for i, v in enumerate(self.sonda_checks):
                    if v.get():
                        mascara |= (1 << i)
                cmd = f"ESTA#{modo}#{intervalo:06d}#{mascara:03d}#"
                self.after(0, self._log, f"-> {cmd}")
                resp = self.conn.comando(cmd)
                self.after(0, self._log, f"<- {resp}")
                if "OK" in resp:
                    self.after(0, messagebox.showinfo,
                               "Estado enviado",
                               f"Comando enviado correctamente:\n{cmd}")
                else:
                    self.after(0, messagebox.showerror,
                               "Error", f"Respuesta inesperada:\n{resp}")
            except Exception as e:
                self.after(0, self._log, f"Error enviando estado: {e}")
        threading.Thread(target=hilo, daemon=True).start()

    def _enviar_idioma(self):
        if not self.conn.conectado:
            messagebox.showwarning("Sin conexion","Conecta primero."); return
        def hilo():
            try:
                idioma = int(self.idioma_var.get().split(" ")[0])
                cmd = f"IDIO#{idioma}#"
                self.after(0, self._log, f"-> {cmd}")
                resp = self.conn.comando(cmd)
                self.after(0, self._log, f"<- {resp}")
                if "OK" in resp:
                    self.after(0, messagebox.showinfo,
                               "Idioma cambiado",
                               f"Idioma enviado correctamente:\n{cmd}")
                else:
                    self.after(0, messagebox.showerror,
                               "Error", f"Respuesta inesperada:\n{resp}")
            except Exception as e:
                self.after(0, self._log, f"Error enviando idioma: {e}")
        threading.Thread(target=hilo, daemon=True).start()
    def _enviar_nueva_ip(self):
        if not self.conn.conectado:
            messagebox.showwarning("Sin conexion","Conecta primero."); return
        nueva = self.nueva_ip_var.get().strip()
        if not nueva:
            messagebox.showwarning("Sin IP","Escribe la nueva IP primero."); return
        def hilo():
            try:
                partes = nueva.split(".")
                if len(partes) != 4:
                    self.after(0, messagebox.showerror,
                               "Error", "IP no valida. Formato: 192.168.0.251")
                    return
                cmd = f"EIP#{int(partes[0]):03d}#{int(partes[1]):03d}#{int(partes[2]):03d}#{int(partes[3]):03d}#"
                self.after(0, self._log, f"-> {cmd}")
                resp = self.conn.comando(cmd)
                self.after(0, self._log, f"<- {resp}")
                p = resp.strip().split(";")
                if p[0] == "EIP" and "OK" in resp:
                    self.after(0, self.ip_var.set, nueva)
                    self.after(0, self.nueva_ip_var.set, "")
                    if self.modo_var.get() == "tcp":
                        self.after(0, self._log, "Reconectando a nueva IP...")
                        self.after(500, self._reconectar_nueva_ip, nueva)
                    else:
                        self.after(0, messagebox.showinfo,
                                   "IP cambiada",
                                   f"IP cambiada correctamente a: {nueva}")
                else:
                    self.after(0, messagebox.showerror,
                               "Error", f"Respuesta inesperada:\n{resp}")
            except Exception as e:
                self.after(0, self._log, f"Error cambiando IP: {e}")
        threading.Thread(target=hilo, daemon=True).start()

    def _reconectar_nueva_ip(self, nueva_ip):
        try:
            self.conn.desconectar()
            time.sleep(1)
            self.conn.conectar_tcp(nueva_ip, self.pto_var.get())
            self._log(f"Reconectado a nueva IP: {nueva_ip}")
            self.led_var.set("  CONECTADO")
            self.led_lbl.config(fg=COLOR_OK)
            self._leer_estado()
            self._leer_version()
            self._leer_uuid()
            self._leer_macs()
            self.after(500, self._actualizar_estado_entries_mac)
        except Exception as e:
            self._log(f"Error reconectando a {nueva_ip}: {e}")
            self.led_var.set("  DESCONECTADO")
            self.led_lbl.config(fg=COLOR_ERROR)
            self.btn_con.config(state="normal")
            self.btn_des.config(state="disabled")

    def _leer_sondas(self):
        if not self.conn.conectado:
            messagebox.showwarning("Sin conexion","Conecta primero."); return
        def hilo():
            try:
                resp = self.conn.comando("?LSO")
                self.after(0, self._log, f"<- {resp}")
                vals = parsear_LSO(resp)
                if vals:
                    self.after(0, self._actualizar_sondas, vals)
                else:
                    self.after(0, self._log, f"Respuesta inesperada: {resp}")
            except Exception as e:
                self.after(0, self._log, f"Error leyendo sondas: {e}")
        threading.Thread(target=hilo, daemon=True).start()

    def _actualizar_sondas(self, vals):
        for i, v in enumerate(vals):
            try:
                num = float(v)
                if num <= -99:
                    self.sonda_vars[i].set("  ---  ")
                else:
                    self.sonda_vars[i].set(f"{num:6.2f}")
            except:
                self.sonda_vars[i].set("  ERR  ")

    def _toggle_continua(self):
        if self.lec_continua_var.get():
            self.leyendo_continuo = True
            self._bucle_continuo()
        else:
            self.leyendo_continuo = False

    def _bucle_continuo(self):
        if not self.leyendo_continuo or not self.conn.conectado:
            self.lec_continua_var.set(False)
            self.leyendo_continuo = False
            return
        self._leer_sondas()
        try:
            intervalo = max(1, int(self.intervalo_lec_var.get())) * 1000
        except:
            intervalo = 2000
        self.after(intervalo, self._bucle_continuo)

    # --------------------------------------------------------------------------
    #   DESCARGA SD
    # --------------------------------------------------------------------------
    def _descargar_todo(self):
        if self.filas_sd == 0:
            self._leer_estado()
        if self.filas_sd == 0:
            messagebox.showinfo("Sin datos","No hay lineas grabadas."); return
        self.desde_var.set("1")
        self.hasta_var.set(str(self.filas_sd))
        self._iniciar_descarga(1, self.filas_sd)

    def _descargar_rango(self):
        try:
            d = int(self.desde_var.get())
            h = int(self.hasta_var.get()) if self.hasta_var.get() else self.filas_sd
            if d < 1 or h < d:
                messagebox.showerror("Error","Rango no valido."); return
            self._iniciar_descarga(d, h)
        except ValueError:
            messagebox.showerror("Error","Valores numericos invalidos.")

    def _iniciar_descarga(self, desde, hasta):
        if not self.conn.conectado:
            messagebox.showwarning("Sin conexion","Conecta primero."); return
        if self.descargando: return
        self.descargando = True
        self.btn_parar.config(state="normal")
        threading.Thread(target=self._hilo_descarga,
                         args=(desde,hasta), daemon=True).start()

    def _hilo_descarga(self, desde, hasta):
        total = hasta - desde + 1
        errores = 0
        self.after(0, self._log,
                   f"Descargando {desde}-{hasta} ({total} filas)...")
        for i, n in enumerate(range(desde, hasta+1)):
            if not self.descargando: break
            try:
                resp = self.conn.comando(f"?LFI#{n}#")
                fila = parsear_LFI(resp, n)
                if fila:
                    self.after(0, self._insertar_fila, fila)
                else:
                    errores += 1
                    self.after(0, self._log, f"Error linea {n}: {resp}")
                pct = ((i+1)/total)*100
                self.after(0, self._prog, pct, i+1, total, errores)
            except Exception as e:
                errores += 1
                self.after(0, self._log, f"Excepcion linea {n}: {e}")
                time.sleep(0.1)
        self.descargando = False
        time.sleep(0.3)
        try:
            if self.conn.modo == "tcp":
                self.conn.sock.settimeout(0.2)
                try:
                    while self.conn.sock.recv(1024): pass
                except Exception:
                    pass
                self.conn.sock.settimeout(3)
            elif self.conn.modo == "serie":
                self.conn.ser.reset_input_buffer()
        except Exception:
            pass
        self.after(0, self.btn_parar.config, {"state":"disabled"})
        self.after(0, self._log,
                   f"Completo. {total-errores} OK, {errores} errores.")

    def _parar_descarga(self):
        self.descargando = False
        self.btn_parar.config(state="disabled")
        self._log("Descarga detenida.")

    def _prog(self, pct, act, tot, err):
        self.prog_var.set(pct)
        self.prog_lbl.config(text=f"{act}/{tot}  ({pct:.0f}%)  Errores:{err}")

    def _insertar_fila(self, fila):
        self.datos.append(fila)
        n = len(self.datos)
        err = any(v.strip().startswith("-99") for v in fila[7:14])
        tag = "err" if err else ("par" if n%2==0 else "impar")
        self.tree.insert("","end", values=fila, tags=(tag,))
        self.tree.yview_moveto(1.0)

    def _limpiar_tabla(self):
        self.tree.delete(*self.tree.get_children())
        self.datos.clear()
        self.prog_var.set(0)
        self.prog_lbl.config(text="")
        self._log("Tabla limpiada.")

    def _ver_grafica(self):
        if not self.datos:
            messagebox.showinfo("Sin datos","No hay datos para graficar."); return

        import datetime

        # Recoger datos
        tiempos = []
        temps   = [[] for _ in range(7)]

        for fila in self.datos:
            try:
                # fila: [linea, fecha(DD/MM/YY), hora(HH:MM:SS), T1..T7, E1..E7]
                dt = datetime.datetime.strptime(
                    f"{fila[1]} {fila[2]}", "%d/%m/%y %H:%M:%S")
                tiempos.append(dt)
                for i in range(7):
                    val = float(fila[3 + i])
                    temps[i].append(None if val <= -99 else val)
            except:
                pass

        if not tiempos:
            messagebox.showinfo("Error","No se pudieron parsear los datos."); return

        # Ventana grafica
        win = tk.Toplevel(self)
        win.title("Grafica Temperatura vs Tiempo")
        win.configure(bg=COLOR_FONDO)
        win.geometry("1100x600")

        fig, ax = plt.subplots(figsize=(11, 5))
        fig.patch.set_facecolor("#1e1e2e")
        ax.set_facecolor("#2a2a3e")
        ax.tick_params(colors="white")
        ax.xaxis.label.set_color("white")
        ax.yaxis.label.set_color("white")
        ax.title.set_color("#7aa2f7")
        for spine in ax.spines.values():
            spine.set_edgecolor("#6c7086")

        colores = ["#7aa2f7","#9ece6a","#f7768e","#e0af68",
                   "#bb9af7","#2ac3de","#ff9e64"]
        nombres = ["Referencia","Sonda 1","Sonda 2","Sonda 3",
                   "Sonda 4","Sonda 5","Sonda 6"]

        for i in range(7):
            # Filtrar None para no romper la linea
            tx = [t for t,v in zip(tiempos, temps[i]) if v is not None]
            ty = [v for v in temps[i] if v is not None]
            if ty:
                ax.plot(tx, ty, label=nombres[i],
                        color=colores[i], linewidth=1.5)

        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
        fig.autofmt_xdate()
        ax.set_xlabel("Tiempo", color="white")
        ax.set_ylabel("Temperatura (°C)", color="white")
        ax.set_title("Temperatura vs Tiempo", color="#7aa2f7")
        ax.grid(True, color="#313244", linestyle="--", alpha=0.5)
        ax.legend(facecolor="#2a2a3e", edgecolor="#6c7086",
                  labelcolor="white", fontsize=8)

        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=8, pady=8)

        # Barra de herramientas
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, win)
        # Anotacion flotante
        annot = ax.annotate("", xy=(0,0), xytext=(15,15),
                            textcoords="offset points",
                            bbox=dict(boxstyle="round", fc="#2a2a3e",
                                      ec="#7aa2f7", alpha=0.9),
                            arrowprops=dict(arrowstyle="->",
                                           color="#7aa2f7"),
                            color="white", fontsize=9)
        annot.set_visible(False)

        # Linea vertical cursor
        vline = ax.axvline(x=tiempos[0], color="#6c7086",
                           linestyle="--", linewidth=1, visible=False)

        def on_move(event):
            if event.inaxes != ax:
                annot.set_visible(False)
                vline.set_visible(False)
                canvas.draw_idle()
                return

            cursor_x = event.xdata
            if cursor_x is None: return

            # Buscar el indice de tiempo mas cercano
            xdata_num = mdates.date2num(tiempos)
            dists = [abs(x - cursor_x) for x in xdata_num]
            idx = dists.index(min(dists))

            # Construir texto con todas las sondas en ese instante
            hora_txt = tiempos[idx].strftime("%H:%M:%S")
            lineas = [f"  {hora_txt}  "]
            for i in range(7):
                val = temps[i][idx] if idx < len(temps[i]) else None
                if val is not None:
                    lineas.append(f"  {nombres[i]}: {val:.2f} °C  ")
                    
            txt = "\n".join(lineas)

            # Posicionar la anotacion en la primera sonda valida
            y_ref = next((temps[i][idx] for i in range(7)
                         if idx < len(temps[i]) and
                         temps[i][idx] is not None), event.ydata)

            annot.xy = (xdata_num[idx], y_ref)
            annot.set_text(txt)
            annot.set_visible(True)
            vline.set_xdata([xdata_num[idx]])
            vline.set_visible(True)
            canvas.draw_idle()

        fig.canvas.mpl_connect("motion_notify_event", on_move)
        toolbar.update()

    def _exportar_csv(self):
        if not self.datos:
            messagebox.showinfo("Sin datos","Nada que exportar."); return
        ruta = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV separado por ;","*.csv"),("Todos","*.*")],
            initialfile="Calorimetro_Log.csv")
        if not ruta: return
        try:
            with open(ruta,"w",newline="",encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(CABECERAS)
                w.writerows(self.datos)
            n = len(self.datos)
            self._log(f"Exportado: {os.path.basename(ruta)} ({n} filas)")
            messagebox.showinfo("OK",f"Guardado:\n{ruta}\n{n} filas.")
        except Exception as e:
            messagebox.showerror("Error al exportar", str(e))

# ==============================================================================
if __name__ == "__main__":
    App().mainloop()
