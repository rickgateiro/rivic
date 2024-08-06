from flask import Flask, render_template, request, send_file
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

@app.route('/download', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Verifica se o arquivo foi enviado pelo formulário.
        file = request.files['file']

        # Verificar se o usuário escolheu um arquivo ou cancelou a seleção
        if not file:
            return "Nenhum arquivo selecionado. Encerrando o programa."



        # Salvar o arquivo Excel
        workbook.save(f'{primeira_celula}_PENDENCIAS.xlsx')
        return send_file(f'{primeira_celula}_PENDENCIAS.xlsx', as_attachment=True)
    return render_template('index.html')  # Crie a página "index.html" no diretório "templates".
if __name__ == '__main__':
    app.run(debug=True)