import time
from customtkinter import *
from tkinter import messagebox
from threading import Thread
from src.database.conexion import ejecutar_query
from src.utils.resource_path import resource_path
from src.models.generar_cartas import generar_cartas
from src.models.seleccionar_archivos import seleccionar_dacxanalista, seleccionar_dac_cdr


class App():    
    def deshabilitar_botones(self):
        self.boton_ejecutar.configure(state="disabled")
        self.boton_dacx.configure(state="disabled")
        self.boton_dac_cdr.configure(state="disabled")
    
    def habilitar_botones(self):
        self.boton_ejecutar.configure(state="normal")
        self.boton_dacx.configure(state="normal")
        self.boton_dac_cdr.configure(state="normal")
    
    def verificar_thread(self, thread):
        if thread.is_alive():
            self.app.after(1000, self.verificar_thread, thread)
        else:
            self.habilitar_botones()
    
    def iniciar_tarea(self):
        self.deshabilitar_botones()
        thread = Thread(target=self.ejecutar_tarea)
        thread.start()
        self.app.after(1000, self.verificar_thread, thread)
    
    def ejecutar_tarea(self):
        self.progressbar.start()
        self.cuadro.configure(state="normal")
        self.cuadro.delete("1.0", "end")
        query = """SELECT * FROM RUTAS WHERE ID == 0"""
        try:
            datos = ejecutar_query(query)
            ruta_dacxa = datos[0][1]
            ruta_dac_cdr = datos[0][2]
            if ruta_dacxa is None or ruta_dac_cdr is None:
                messagebox.showerror("Error", "Por favor, configure las rutas de los archivos.")
            elif not os.path.exists(ruta_dacxa):
                messagebox.showerror("Error", "No se encontraró el archivo DACxANALISTA en la ruta especificada.")
            elif not os.path.exists(ruta_dac_cdr):
                messagebox.showerror("Error", "No se encontraró el archivo DAC y CDR en la ruta especificada.")
            else:
                start = time.time()
                generar_cartas(ruta_dacxa, ruta_dac_cdr, self.cuadro)
        except Exception as ex:
            messagebox.showerror("Error", "Detalle:\n" + str(ex))
        finally:
            end = time.time()
            self.progressbar.stop()
            if start is not None:
                tiempo_promedio = end - start
                self.cuadro.insert("end", f"Tiempo de ejecución: {(round(tiempo_promedio, 2))} segundos.\n")
            else:
                self.cuadro.insert("end", "No se ejecutó la tarea.\n")
            self.cuadro.configure(state="disabled")
    
    def crear_app(self):
        self.app = CTk()
        self.app.title("Cartas de Requerimiento de Pago")
        icon_path = resource_path("./src/images/icono.ico")
        if os.path.isfile(icon_path):
            self.app.iconbitmap(icon_path)
        else:
            messagebox.showwarning("ADVERTENCIA", "No se encontró el archivo 'icono.ico' en la ruta: " + icon_path)
        self.app.resizable(False, False)
        set_appearance_mode("dark")
        
        main_frame = CTkFrame(self.app)
        main_frame.pack_propagate("True")
        main_frame.pack(fill="both", expand=True)
        
        frame_base = CTkFrame(main_frame)
        frame_base.grid(row=0, column=0, padx=(20, 10), pady=(20, 0), sticky="nsew")
        
        ruta_dacxa = CTkLabel(frame_base, text="Ruta DACxAnalista", font=("Calibri",15))
        ruta_dacxa.pack(padx=(20, 20), pady=(5, 0), fill="both", expand=True, anchor="center", side="top")
        self.boton_dacx = CTkButton(frame_base, text="Seleccionar", font=("Calibri",15), text_color="white",
                                fg_color="transparent", border_color="#d11515", border_width=2, hover_color="#d11515", 
                                width=25, corner_radius=25, command=lambda: seleccionar_dacxanalista())
        self.boton_dacx.pack(padx=(20, 20), pady=(0, 15), fill="both", anchor="center", side="bottom")
        
        frame_dacx = CTkFrame(main_frame)
        frame_dacx.grid(row=0, column=1, padx=(10, 20), pady=(20, 0), sticky="nsew")
        
        ruta_daccdr = CTkLabel(frame_dacx, text="Ruta DAC y CDR", font=("Calibri",15))
        ruta_daccdr.pack(padx=(20, 20), pady=(5, 0), fill="both", expand=True, anchor="center", side="top")
        self.boton_dac_cdr = CTkButton(frame_dacx, text="Seleccionar", font=("Calibri",15), text_color="white",
                                fg_color="transparent", border_color="#d11515", border_width=2, hover_color="#d11515", 
                                width=25, corner_radius=25, command=lambda: seleccionar_dac_cdr())
        self.boton_dac_cdr.pack(padx=(20, 20), pady=(0, 15), fill="both", anchor="center", side="bottom")
        
        self.boton_ejecutar = CTkButton(main_frame, text="GENERAR CARTAS", text_color="black", font=("Calibri",20,"bold"), 
                                    border_color="black", border_width=3, fg_color="gray", 
                                    hover_color="red", command=lambda: self.iniciar_tarea())
        self.boton_ejecutar.grid(row=1, column=0, columnspan=2, ipady=20, padx=(20, 20), pady=(20, 0), sticky="nsew")
        
        self.cuadro = CTkTextbox(main_frame, font=("Calibri",15), height=110, border_color="black", border_width=2)
        self.cuadro.grid(row=2, column=0, columnspan=2, padx=(20, 20), pady=(20, 0), sticky="nsew")
        self.cuadro.configure(state="disabled")
        
        self.progressbar = CTkProgressBar(main_frame, mode="indeterminate", orientation="horizontal", 
                                        progress_color="#d11515", height=10, border_width=0)
        self.progressbar.grid(row=3, column=0, columnspan=2, padx=(20, 20), pady=(20, 20), sticky="nsew")
        
        self.app.mainloop()