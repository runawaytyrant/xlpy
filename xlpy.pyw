import pandas as pd
import os
from tkinter import Tk, filedialog, Button, Text, Label, Entry, Frame

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xls*; *.xlsx; *.xlsm; *.xlsb")])
    if file_path:
        update_label(file_path)
        possible_column_combinations = [
            ['Matrícula', 'A. Paterno', 'A. Materno', 'Nombre(s)']
        ]
        for cols in possible_column_combinations:
            try:
                df = pd.read_excel(file_path, usecols=cols)
                df['Matrícula'] = df['Matrícula'].astype(str)
                return df
            except (KeyError, ValueError):
                continue
        text_widget.insert('1.0', "Error. Seleccione otro archivo.\n\n")
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
    
    if starting_number and df is not None:
        filtered_df = df[df['Matrícula'].str.startswith(starting_number)]
        if not filtered_df.empty:
            text = filtered_df.to_string(index=False)
            count_text = "\n\nTotal: " + str(len(filtered_df))
            text_widget.insert('1.0', text + count_text + "\n\n")
        else:
            text_widget.insert('1.0', "0 coincidencias encontradas.\n\n")
    elif starting_number == "":
        text_widget.insert('1.0', "Introduzca un número.\n\n")

if __name__ == "__main__":
    root = Tk()
    root.title("XlPy")
    root.geometry("465x420")

    def browse_and_load_data():
        global df
        df = browse_file()

    def filter_data():
        if df is not None:
            generate_output(text_widget, starting_number_entry, df)

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

    #graph_button = Button(filtering_frame, text="Generar gráfica")
    #graph_button.grid(row=0, column=3, padx=40)

    root.mainloop()
