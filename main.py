#!/usr/bin/env python3

import tkinter as tk
import tkinter.ttk as ttk
from appdirs import user_config_dir
from pathlib import Path
from os.path import join as path_join
import configparser
from os.path import exists as path_exists
from os.path import isfile as file_exists
from tkinter import filedialog as fd
from tkinter.messagebox import showerror, showinfo
from shutil import move as copy_file
from openpyxl import load_workbook
from fillpdf import fillpdfs
from num2words import num2words
import os
import base64

def empty_validation(input):
    if not input or input == '' or input == None:
        return 0
    else:
        return 1

def empty_message_error():
    showerror(
        title='Error de validación',
        message='Los campos no pueden estar vacíos.'
    )

app_name = 'IPSPAportes'
author_name = 'Daniel_Jimenez'
file_config_name = 'ipsp_aportes.ini'
file_pdf_name = 'PLANTILLA_FORMULARIO.pdf'
file_empty_message = 'Debe seleccionar un archivo'
path_empty_message = 'Debe seleccionar un directorio'
app_config_path = user_config_dir(app_name, author_name)
Path(app_config_path).mkdir(parents=True, exist_ok=True)
file_config_path = path_join(app_config_path, file_config_name)
file_pdf_path = path_join(app_config_path, file_pdf_name)

config = configparser.ConfigParser()
if path_exists(file_config_path):
    config.read(file_config_path)
if not config.has_section('INPUT'):
    config.add_section('INPUT')
if not config.has_section('SIGNER'):
    config.add_section('SIGNER')

root = tk.Tk()
s = ttk.Style(root)
s.theme_use('default')
s.configure('TNotebook.Tab', font=('Arial', '12', 'normal'))
s.configure('custom_button.TButton', font=('Arial', '12', 'normal'))

root.title('MAS IPSP - APORTES')
empty_validation_command = root.register(empty_validation)
empty_message_error_command = root.register(empty_message_error)

notebook = ttk.Notebook(root)
notebook.pack(fill=tk.BOTH, expand=True)
notebook.pressed_index = None

tab1 = ttk.Frame(notebook)
tab1.pack(fill=tk.BOTH, expand=True)

tab2 = ttk.Frame(notebook)
tab2.pack(fill=tk.BOTH, expand=True)

img1 = tk.PhotoImage(data='iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAMAAAAoLQ9TAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAABv1BMVEUAAABGRkEqKiciIiAkJCJISEMODg4LCwtISEQFBQUCAgIDAwMDAwMDAwMDAwMDAwIDAwMCAgICAgIgIB4DAwMDAwMODg0DAwMLCwoPDw4ICAgCAgIBAQEBAQEHBwYDAwMEBAQSEhEBAQEBAQEHBwYMDAsEBAQEBAQEBAQGBgUQEBACAgEBAQECAgIICAcJCQkCAgICAgEFBQUGBgUGBgUHBwYFBQUBAQECAgEDAwNbW1UEBAMCAgIBAQEBAQEBAQEGBgYGBgYGBgYJCQgCAgICAgIKCgl4eHEDAwMCAgICAgICAgICAgIFBQQFBQUDAwMBAQFAQD1HR0NEREAEBAQVFRQNDQwQEA8SEhEYGBcJCQkBAQEDAwNVVVBEREAEBAMEBAQDAwMDAwMCAgIDAwM1NTINDQ0ODg0BAQEBAQEEBAQ2NjICAgIDAwMDAwMQEA8qKig3NzMdHRsJCQg2NjMEBAMEBAQDAwMEBAQiIiBhYVsCAgIDAwICAgIDAwJpaWNKSkVUVE+KioL///8KCgoCAgJTU08CAgIDAwMDAwMDAwMiIiFYWFMPDw8GBgYGBgUGBgUICAhLS0hKSkb///8SAPvFAAAAlHRSTlMAAwYHBgMMFQMtb2RiYWVrVY+ZB0ZRE2AZEh+B8b0jRUQMxtwcFjw9Pi8QsuiCHhykszEtLiQvybBLAkym5NjLKi0pFKCSFwJOhKOdmDkyXPYDAwRGDRURDwoQ4EwDA0k/VF+MXgUUFLfTNAWvTWMQBgUJHgVKQFZCCANwaZdnAwQDAgEZcwN5WFpZBQMNKCorGgIDCz5HMwAAAAFiS0dEg/y0z9IAAAAHdElNRQflDAwBKxxmyLk5AAAA8klEQVQY02NgZGKGARZWNnYOBk4ubh5ebj4g4OYXEBQSZhARFRMQlxCTBAIpaRlZOQZ5BUUlZRVVNXUNTS1tHV09oIC+gaGRpLGxiamZuYWlFYO8tY2tnb2Do6aTs4urpZs7g7yHp5e3j4mvmp9/gHRgUDCDfEhoWJhYeERklGx0TGxwHIN8fIJAYmKigEZSQHJKKkggJDQtPT1dTDIjM4s3G65CIFEtJ9ctLxuqIi2/oDDcsaiYuwQkUFpWXl5RWSVWXVNbVw8SUGBoaGhsam5p5eEObwMKJLR3dAJBF59PRXdPMFCgt68fDCZMnDQ5DggAHkREcJfzK3gAAAAldEVYdGRhdGU6Y3JlYXRlADIwMjEtMTItMTJUMDE6NDM6MjUrMDA6MDCcBOh7AAAAJXRFWHRkYXRlOm1vZGlmeQAyMDIxLTEyLTEyVDAxOjQzOjI1KzAwOjAw7VlQxwAAACB0RVh0c29mdHdhcmUAaHR0cHM6Ly9pbWFnZW1hZ2ljay5vcme8zx2dAAAAGHRFWHRUaHVtYjo6RG9jdW1lbnQ6OlBhZ2VzADGn/7svAAAAGHRFWHRUaHVtYjo6SW1hZ2U6OkhlaWdodAAxOTJAXXFVAAAAF3RFWHRUaHVtYjo6SW1hZ2U6OldpZHRoADE5MtOsIQgAAAAZdEVYdFRodW1iOjpNaW1ldHlwZQBpbWFnZS9wbmc/slZOAAAAF3RFWHRUaHVtYjo6TVRpbWUAMTYzOTI3MzQwNYMms/UAAAAPdEVYdFRodW1iOjpTaXplADBCQpSiPuwAAABWdEVYdFRodW1iOjpVUkkAZmlsZTovLy9tbnRsb2cvZmF2aWNvbnMvMjAyMS0xMi0xMi8wODEyMTA5YmFjYzExYjAwMGRlM2ZmNWQ5NzUwNjhkNy5pY28ucG5ncXv8dQAAAABJRU5ErkJggg==')
notebook.add(tab1, text='Generación de Formularios', image=img1, compound='left')

img2 = tk.PhotoImage(data='iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAQAAAC1+jfqAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAAAmJLR0QA/4ePzL8AAAAJcEhZcwAAEHMAABBzARg5fEAAAAAHdElNRQflDAwBKwabqkBDAAABMElEQVQoz33Rv0vUARzG8dfd9zhNJQK/XYg/sLZcTq0hp7DJRfwD3LIil6KmaHFwjRoTJ1enliKhRsE4wQZFDkFFMI8MakjIQb+fhsu7IfG9Ps/z4fnwJJpc0a/VkfAfeYl2C75ZcwuFppTgspcm3XBfSZeiQQ9s+940zTgRToUQMpmw7FpdLMi5KgFVValhbUi1nuXbPJcJC/okOjz0U/isV65u+OiXsOV6o9WcsOaZewYgE8IHxUanR2oeW7FtXDGvJpDqaBh6VZXcsenAPEPeCsee/Pv+tlUvLAlHajYLvtpHi1lDKrqNee+msswlfyzm5M2bapz/7Y1+nQ6NeOWLnUTYMuBYRZ/wWo+yDcOeeueHk3ou1aOkYte0PUvWjTqHQRPKPll21wWkZyvU+Qt4YldxF3VvXAAAACV0RVh0ZGF0ZTpjcmVhdGUAMjAyMS0xMi0xMlQwMTo0Mjo1MiswMDowMLyktNIAAAAldEVYdGRhdGU6bW9kaWZ5ADIwMjEtMTItMTJUMDE6NDI6NTIrMDA6MDDN+QxuAAAAIHRFWHRzb2Z0d2FyZQBodHRwczovL2ltYWdlbWFnaWNrLm9yZ7zPHZ0AAAAYdEVYdFRodW1iOjpEb2N1bWVudDo6UGFnZXMAMaf/uy8AAAAYdEVYdFRodW1iOjpJbWFnZTo6SGVpZ2h0ADE5MkBdcVUAAAAXdEVYdFRodW1iOjpJbWFnZTo6V2lkdGgAMTky06whCAAAABl0RVh0VGh1bWI6Ok1pbWV0eXBlAGltYWdlL3BuZz+yVk4AAAAXdEVYdFRodW1iOjpNVGltZQAxNjM5MjczMzcyV0ymFAAAAA90RVh0VGh1bWI6OlNpemUAMEJClKI+7AAAAFZ0RVh0VGh1bWI6OlVSSQBmaWxlOi8vL21udGxvZy9mYXZpY29ucy8yMDIxLTEyLTEyLzNjOTliMDYxNmI0OWFkMmNkODUyNjNmMTEyNTA5NmM1Lmljby5wbmcNsPFfAAAAAElFTkSuQmCC')
notebook.add(tab2, text='Configuración', image=img2, compound='left')

tab1.wb = None
tab1.ws = None

input_sheet = tk.StringVar(tab1, '')

def load_excel_data():
    if config.has_option('INPUT', 'path'):
        if config['INPUT']['path'] != file_empty_message:
            if file_exists(config['INPUT']['path']):
                tab1.wb = load_workbook(config['INPUT']['path'])
                if len(tab1.wb.sheetnames) > 0:
                    input_sheet.set(tab1.wb.sheetnames[0])
                    tab1.ws = tab1.wb[tab1.wb.sheetnames[0]]
                    return True
    return False

pdf_form_path = tk.StringVar(root, config['FORM']['path'] if config.has_option('FORM', 'path') else file_empty_message)
output_path = tk.StringVar(root, config['OUTPUT']['path'] if config.has_option('OUTPUT', 'path') else path_empty_message)
excel_input_path = tk.StringVar(root, config['INPUT']['path'] if config.has_option('INPUT', 'path') else file_empty_message)

tab1.row_from = tk.IntVar(tab1, 2)
tab1.row_to = tk.IntVar(tab1, 2)
tab1.progress_total = tk.IntVar(tab1, 0)
tab1.progress_current = tk.IntVar(tab1, 0)
tab1.progress_success = tk.IntVar(tab1, 0)
tab1.progress = tk.StringVar(tab1, '0/0')

def select_file(file_types, config_section):
    filename = fd.askopenfilename(
        title='Abrir archivo',
        initialdir=str(Path.home()),
        filetypes=file_types
    )
    if filename:
        showinfo(
            title='Archivo seleccionado',
            message=filename
        )
        if not config.has_section(config_section):
            config.add_section(config_section)
        if config_section == 'FORM':
            pdf_form_path.set(filename)
        if config_section == 'INPUT':
            excel_input_path.set(filename)
        return True
    else:
        return False

def select_folder():
    pathname = fd.askdirectory(
        title='Seleccionar carpeta',
        initialdir=str(Path.home())
    )
    if pathname:
        showinfo(
            title='Carpeta seleccionada',
            message=pathname
        )
        if not config.has_section('OUTPUT'):
            config.add_section('OUTPUT')
        config.set('OUTPUT', 'path', pathname)
        output_path.set(pathname)
        save_config(False)
        return True
    else:
        return False

def select_excel_input_file():
    if (select_file([('Excel', '*.xlsx'), ('Excel', '*.XLSX'), ('Excel', '*.xls'), ('Excel', '*.XLS')], 'INPUT') == True):
        config.set('INPUT', 'path', excel_input_path.get())
        save_config(False)
    else:
        excel_input_path.set(file_empty_message)
        showerror(
            title='Error en la configuración',
            message='Corrija los datos de la configuración.'
        )
        notebook.select(1)

def fill_pdf_template(data):
    data_dict = fillpdfs.get_form_fields(Path(pdf_form_path.get()))
    for i in [1, 2]:
        data_dict['number{}'.format(i)] = f'{int(data[0]):07}'
        data_dict['date_issue{}'.format(i)] = data[1]
        data_dict['date_deposit{}'.format(i)] = data[2]
        data_dict['name{}'.format(i)] = data[3]
        data_dict['ci{}'.format(i)] = data[4]
        data_dict['cel{}'.format(i)] = int(data[5])
        data_dict['militant{}'.format(i)] = 'SI' if data[6] == 'SI' else 'Off'
        data_dict['monthly{}'.format(i)] = 'SI' if data[7] == 'SI' else 'Off'
        data_dict['education{}'.format(i)] = 'SI' if data[8] == 'SI' else 'Off'
        data_dict['money_float{}'.format(i)] = data[9]
        data_dict['money_literal{}'.format(i)] = '{0}{1}/100 BOLIVIANOS'.format(num2words(data_dict['money_float{}'.format(i)], lang='es', to='currency').split('euros')[0].upper(), f'{int(round(float(data[9])%1, 2)*100):02}')
        data_dict['signer_name{}'.format(i)] = config['SIGNER']['name'].upper()
        data_dict['signer_charge{}'.format(i)] = config['SIGNER']['charge'].upper()
        out_file = path_join(Path(config['OUTPUT']['path']), '{}_{}.pdf'.format(data_dict['number1'], data_dict['name1'].replace(' ', '_')))
        if os.path.exists(out_file):
            os.remove(out_file)
        fillpdfs.write_fillable_pdf(Path(pdf_form_path.get()), out_file, data_dict, flatten=False)

def generate_pdfs():
    if not path_exists(config['OUTPUT']['path']):
        output_path.set(path_empty_message)
        showerror(
            title='Error en el carpeta de destino',
            message='La carpeta de destino seleccionado no existe, seleccione nuevamente.'
        )
        return None
    try:
        row_from = tab1.row_from.get()
        row_to = tab1.row_to.get()
    except:
        showerror(
            title='Error en la selección de filas',
            message='El valor de la selección de filas debe ser un número de entre las filas de la Hoja de Excel.'
        )
        return None
    if row_to < 2 or row_from < 2 or row_to > 1048576 or row_from > 1048576 or type(row_to) is not int or type(row_from) is not int:
        showerror(
            title='Error en la selección de filas',
            message='El valor de la selección de filas debe estar entre 2 y 1048576.'
        )
        return None
    if row_to < row_from:
        showerror(
            title='Error en la selección de filas',
            message='El valor Hasta fila debe ser mayor o igual que el valor Desde fila.'
        )
        return None
    if not load_excel_data():
        excel_input_path.set(file_empty_message)
        showerror(
            title='Error al cargar la hoja de cálculo Excel',
            message='El archivo no existe o no corresponde a la plantilla de datos.'
        )
        return None
    tab1.progress_success.set(0)
    tab1.progress_current.set(0)
    tab1.progress_total.set(row_to - row_from + 1)
    button_file['state'] = 'disabled'
    button_folder['state'] = 'disabled'
    entry_row_from['state'] = 'disabled'
    entry_row_to['state'] = 'disabled'
    button_run['state'] = 'disabled'
    for row in tab1.ws.iter_rows(min_row=row_from, max_col=10, max_row=row_to, values_only=True):
        tab1.progress_current.set(tab1.progress_current.get() + 1)
        tab1.progress.set('{0}/{1}'.format(tab1.progress_current.get(), tab1.progress_total.get()))
        escape = False
        for cell in row:
            if not cell or cell == '' or cell == None:
                escape = True
                break
        if escape:
            continue
        else:
            fill_pdf_template(row)
            tab1.progress_success.set(tab1.progress_success.get() + 1)
        tab1.update()
    button_file['state'] = 'normal'
    button_folder['state'] = 'normal'
    entry_row_from['state'] = 'normal'
    entry_row_to['state'] = 'normal'
    button_run['state'] = 'normal'
    showinfo(
        title='Formularios generados',
        message='El proceso ha finalizado satisfactoriamente.'
    )

# Ventana generación de formularios

ttk.Label(tab1, text='Hoja de cálculo Excel de aportantes', font=('Arial', '14', 'bold underline')).grid(sticky='W', column=0, row=0, padx=10, pady=10, columnspan=4)

ttk.Label(tab1, text='Archivo Excel:', font=('Arial', '12', 'normal'), anchor='e').grid(sticky='WE', column=0, row=1, padx=10, pady=5)
ttk.Label(tab1, text='', font=('Arial', '12', 'normal'), anchor='w', textvariable=excel_input_path, wraplength=340).grid(sticky='WE', column=1, row=1, padx=10, pady=5, columnspan=2)
button_file = ttk.Button(tab1, text='Seleccionar archivo', style='custom_button.TButton', command=select_excel_input_file)
button_file.grid(sticky='W', column=3, row=1, padx=10, pady=5)

ttk.Label(tab1, text='Carpeta donde se exportarán los formularios', font=('Arial', '14', 'bold underline')).grid(sticky='W', column=0, row=2, padx=10, pady=10, columnspan=4)

ttk.Label(tab1, text='Carpeta destino:', font=('Arial', '12', 'normal'), anchor='e').grid(sticky='WE', column=0, row=3, padx=10, pady=5)
ttk.Label(tab1, text='', font=('Arial', '12', 'normal'), anchor='w', textvariable=output_path, wraplength=340).grid(sticky='WE', column=1, row=3, padx=10, pady=5, columnspan=2)
button_folder = ttk.Button(tab1, text='Seleccionar carpeta', style='custom_button.TButton', command=lambda:select_folder())
button_folder.grid(sticky='W', column=3, row=3, padx=10, pady=5)

ttk.Label(tab1, text='Desde fila:', font=('Arial', '12', 'normal'), anchor='e').grid(sticky='WE', column=0, row=4, padx=10, pady=5)
entry_row_from = ttk.Entry(tab1, takefocus=0, textvariable=tab1.row_from)
entry_row_from.grid(sticky='WE', column=1, row=4, padx=10, pady=5)
ttk.Label(tab1, text='Hasta fila:', font=('Arial', '12', 'normal'), anchor='e').grid(sticky='WE', column=2, row=4, padx=10, pady=5)
entry_row_to = ttk.Entry(tab1, takefocus=0, textvariable=tab1.row_to)
entry_row_to.grid(sticky='WE', column=3, row=4, padx=10, pady=5)

ttk.Label(tab1, text='Progreso:', font=('Arial', '12', 'normal'), anchor='e').grid(sticky='WE', column=0, row=5, padx=10, pady=5)
ttk.Label(tab1, text='', font=('Arial', '12', 'normal'), anchor='w', textvariable=tab1.progress).grid(sticky='WE', column=1, row=5, padx=10, pady=5)
ttk.Label(tab1, text='PDFs generados:', font=('Arial', '12', 'normal'), anchor='e').grid(sticky='WE', column=2, row=5, padx=10, pady=5)
ttk.Label(tab1, text='', font=('Arial', '12', 'normal'), anchor='w', textvariable=tab1.progress_success).grid(sticky='WE', column=3, row=5, padx=10, pady=5)

button_run = ttk.Button(tab1, text='Generar formularios', style='custom_button.TButton', command=generate_pdfs)
button_run.grid(sticky='E', column=3, row=6, padx=10, pady=20)

# Ventana configuración

ttk.Label(tab2, text='Datos de la autoridad firmante', font=('Arial', '14', 'bold underline')).grid(sticky='W', column=0, row=0, padx=10, pady=10, columnspan=3)

ttk.Label(tab2, text='Nombre:', font=('Arial', '12', 'normal'), anchor='e').grid(sticky='WE', column=0, row=1, padx=0, pady=5)
signer_name = tk.StringVar(root, config['SIGNER']['name'] if config.has_option('SIGNER', 'name') else '')
ttk.Entry(tab2, takefocus=0, textvariable=signer_name, validate='focusout', validatecommand=(empty_validation_command, '%P'), invalidcommand=(empty_message_error_command)).grid(sticky='WE', column=1, row=1, padx=10, pady=5, columnspan=2)

ttk.Label(tab2, text='Cargo:', font=('Arial', '12', 'normal'), anchor='e').grid(sticky='WE', column=0, row=2, padx=0, pady=5)
signer_charge = tk.StringVar(root, config['SIGNER']['charge'] if config.has_option('SIGNER', 'charge') else '')
ttk.Entry(tab2, takefocus=0, textvariable=signer_charge, validate='focusout', validatecommand=(empty_validation_command, '%P'), invalidcommand=(empty_message_error_command)).grid(sticky='WE', column=1, row=2, padx=10, pady=5, columnspan=2)

ttk.Label(tab2, text='Plantilla de formulario pdf', font=('Arial', '14', 'bold underline')).grid(sticky='W', column=0, row=3, padx=10, pady=10, columnspan=3)

ttk.Label(tab2, text='Formulario PDF:', font=('Arial', '12', 'normal'), anchor='e').grid(sticky='WE', column=0, row=4, padx=10, pady=5)
ttk.Label(tab2, text='', font=('Arial', '12', 'normal'), anchor='w', textvariable=pdf_form_path, wraplength=340).grid(sticky='WE', column=1, row=4, padx=10, pady=5)
ttk.Button(tab2, text='Seleccionar archivo', style='custom_button.TButton', command=lambda:select_file([('pdf file', '*.pdf'), ('pdf file', '*.PDF')], 'FORM')).grid(sticky='W', column=2, row=4, padx=10, pady=5)

def save_config(show_info=True):
    signer_name_value = signer_name.get().strip().upper()
    signer_charge_value = signer_charge.get().strip().upper()
    if not signer_name_value or not signer_charge_value:
        empty_message_error()
        notebook.select(1)
    else:
        config.set('SIGNER', 'name', signer_name_value)
        config.set('SIGNER', 'charge', signer_charge_value)
        if file_pdf_path != pdf_form_path.get():
            copy_file(pdf_form_path.get(), file_pdf_path)
            config.set('FORM', 'path', file_pdf_path)
            pdf_form_path.set(file_pdf_path)
        with open(file_config_path, 'w') as configfile:
            config.write(configfile)
        if show_info == True:
            showinfo(title='Configuración', message='Configuración guardada exitosamente.')
            notebook.select(0)

ttk.Button(tab2, text='Guardar', style='custom_button.TButton', command=save_config).grid(sticky='E', column=2, row=5, padx=10, pady=20)

if path_exists(file_config_path):
    notebook.select(0)
else:
    notebook.select(1)

icon = tk.PhotoImage(data='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')
root.tk.call('wm', 'iconphoto', root._w, icon)

root.geometry('')
root.mainloop()
