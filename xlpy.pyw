import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from tkinter import Tk, filedialog, Button, Text, Label, Entry, Frame

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xls*; *.xlsx; *.xlsm; *.xlsb")])
    update_label(file_path)
    if file_path:
        text_widget.delete('1.0', 'end')
        possible_column_combinations = [
            ['Matrícula', 'A. Paterno', 'A. Materno', 'Nombre(s)'],
            ['CARRERA', 'TOTAL EGRESADOS', 'TOTAL TRABAJAN', 'CONTINUIDAD ESTUDIOS']
        ]
        for cols in possible_column_combinations:
            try:
                df = pd.read_excel(file_path, usecols=cols)
                return df
            except (KeyError, ValueError):
                continue
        text_widget.insert('1.0', "Error: Seleccione otro archivo.")
    return None

def update_label(file_path):
    if file_path:
        filename = os.path.basename(file_path)
        message_label.config(text="Archivo: " + filename)
    else:
        message_label.config(text="Archivo:")

def generate_output(text_widget, starting_number_entry, df):
    pd.options.display.max_rows = 99999

    starting_number = starting_number_entry.get()

    text_widget.delete('1.0', 'end')
    
    if starting_number and df is not None:
        try:
            df['Matrícula'] = df['Matrícula'].astype(str)
            filtered_df = df[df['Matrícula'].str.startswith(starting_number)]
            if not filtered_df.empty:
                text = filtered_df.to_string(index=False)
                count_text = "\n\nTotal: " + str(len(filtered_df))
                text_widget.insert('1.0', text + count_text)
            else:
                text_widget.insert('1.0', "0 coincidencias encontradas.")
        except (KeyError, ValueError):
                text_widget.insert('1.0', "Este archivo no se puede filtrar.")
    elif not starting_number:
                text_widget.insert('1.0', "Introduzca un número.")

def generate_graph(df):
    if df is not None:
        try:
            x = df['CARRERA']
            y1 = df['TOTAL EGRESADOS']
            y2 = df['TOTAL TRABAJAN']
            y3 = df['CONTINUIDAD ESTUDIOS']
            
            y_positions = np.arange(len(x))
            
            width = 0.3  # the width of the bars
            fig, ax = plt.subplots(figsize=(13, len(x) * 0.5))

            ax.barh(y_positions - width, y1, width, label='TOTAL EGRESADOS')
            ax.barh(y_positions, y2, width, label='TOTAL TRABAJAN')
            ax.barh(y_positions + width, y3, width, label='CONTINUIDAD ESTUDIOS')

            ax.set_ylabel('')
            ax.set_xlabel('Cantidad de alumnos')
            ax.set_title('Seguimiento de Egresados')
            ax.set_yticks(y_positions)
            ax.set_yticklabels(x)
            ax.legend()

            plt.tight_layout(pad=3.0)
            plt.show()
        except KeyError:
            text_widget.delete('1.0', 'end')
            text_widget.insert('1.0', "Las columnas necesarias para la gráfica no están presentes en el archivo.")
    else:
        text_widget.delete('1.0', 'end')
        text_widget.insert('1.0', "No se ha cargado ningún archivo.")


if __name__ == "__main__":
    root = Tk()
    root.title("XlPy")
    root.geometry("465x420")

    df = None

    def browse_and_load_data():
        global df
        df = browse_file()

    def filter_data():
        if df is not None:
            generate_output(text_widget, starting_number_entry, df)
        else:
            text_widget.delete('1.0', 'end')
            text_widget.insert('1.0', "No se ha cargado ningún archivo.")

    file_frame = Frame(root)
    file_frame.grid(row=0, column=0, padx=10, pady=5, sticky="w")

    message_label = Label(file_frame, text="Archivo:")
    message_label.grid(row=0, column=0)

    browse_button = Button(file_frame, text="Abrir", command=browse_and_load_data)
    browse_button.grid(row=0, column=1)

    text_widget = Text(root, height=20, width=55)
    text_widget.grid(row=1, column=0, padx=10, pady=5)

    filtering_frame = Frame(root)
    filtering_frame.grid(row=2, column=0, padx=10, pady=5, sticky="w")

    starting_number_label = Label(filtering_frame, text="Número de inicio:")
    starting_number_label.grid(row=0, column=0)

    starting_number_entry = Entry(filtering_frame)
    starting_number_entry.grid(row=0, column=1)

    filter_button = Button(filtering_frame, text="Filtrar", command=filter_data)
    filter_button.grid(row=0, column=2, padx=5)

    graph_button = Button(filtering_frame, text="Generar gráfica", command=lambda:generate_graph(df))
    graph_button.grid(row=0, column=3, padx=10)

    root.mainloop()
