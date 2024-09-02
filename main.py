import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook, Workbook

class ExcelGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Repuestoturbo1000 (hay que buscar un buen nombre)")
        self.master.geometry("500x400")

        self.excel_file = "Repuestos2024.xlsx"
        self.create_excel_if_not_exists()

        # Frame para la referencia
        self.ref_frame = tk.Frame(self.master)
        self.ref_frame.pack(pady=10)

        tk.Label(self.ref_frame, text="Ref.Componente").grid(row=0, column=0)
        self.ref_entry = tk.Entry(self.ref_frame)
        self.ref_entry.grid(row=0, column=1)
        tk.Button(self.ref_frame, text="Buscar", command=self.buscar_datos).grid(row=0, column=2)

        # Frame para mostrar y modificar datos
        self.data_frame = tk.Frame(self.master)
        self.data_frame.pack(pady=10)

        self.labels = ["Frontal", "Lateral Der.", "Lateral Izq.", "Power/Reset", "Leds frontales", "Varios", "Protecciones"]
        self.entries = []

        for i, label in enumerate(self.labels):
            tk.Label(self.data_frame, text=label).grid(row=i, column=0)
            entry = tk.Entry(self.data_frame)
            entry.grid(row=i, column=1)
            self.entries.append(entry)

        # Botón para guardar cambios
        tk.Button(self.master, text="Guardar", command=self.guardar_cambios).pack(pady=10)

    def create_excel_if_not_exists(self):
        try:
            load_workbook(self.excel_file)
        except FileNotFoundError:
            wb = Workbook()
            ws = wb.active
            ws.append(["Frontal", "Lateral Der.", "Lateral Izq.", "Power/Reset", "Leds frontales", "Varios", "Protecciones"])
            wb.save(self.excel_file)

    def buscar_datos(self):
        referencia = self.ref_entry.get()
        if not referencia:
            messagebox.showerror("Error", "Mira ver que has liao y vuelve a intentarlo")
            return

        wb = load_workbook(self.excel_file)
        ws = wb.active

        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] == referencia:
                for entry, value in zip(self.entries, row[1:]):
                    entry.delete(0, tk.END)
                    entry.insert(0, str(value))
                return

        messagebox.showinfo("Nueva Referencia", "No hay Repuestos de esta referencia, Introduce datos para añadir repuestos.")

    def guardar_cambios(self):
        referencia = self.ref_entry.get()
        if not referencia:
            messagebox.showerror("Error", "Mira ver que has liao y vuelve a intentarlo")
            return

        datos = [entry.get() for entry in self.entries]

        wb = load_workbook(self.excel_file)
        ws = wb.active

        row_to_update = None
        for row in ws.iter_rows(min_row=2):
            if row[0].value == referencia:
                row_to_update = row
                break

        if row_to_update:
            for cell, value in zip(row_to_update[1:], datos):
                cell.value = value
        else:
            ws.append([referencia] + datos)

        wb.save(self.excel_file)
        messagebox.showinfo("Involucro", "Datos guardados, a currar")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelGUI(root)
    root.mainloop()