import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.units import inch
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename
import os

def formatar_valores(data):
    """ Remove o .0 dos valores flutuantes e formata os números inteiros """
    formatted_data = []
    for row in data:
        new_row = []
        for cell in row:
            if isinstance(cell, float) and cell.is_integer():
                new_row.append(int(cell))
            else:
                new_row.append(cell)
        formatted_data.append(new_row)
    return formatted_data

def gerar_relatorio_pecas(df, arquivo_excel):
    # Substitui NaN por string vazia
    df = df.fillna('')
    
    # Aplicar a lógica de exclusão de linhas
    mask_exclude = (
        ((df['PEÇA DESCRIÇÃO'].str.contains('_PAINEL_DUP_', na=False)) & (df['ESPESSURA'].isin([15, 18]))) |
        (df['PEÇA DESCRIÇÃO'].str.contains('_PRAT_DUP_CORTE', na=False)) |
        (df['PEÇA DESCRIÇÃO'].str.contains('_PAINEL_ENG_CORTE', na=False)) |
        (df['PEÇA DESCRIÇÃO'].str.contains('_ENGROSSO_', na=False)) |
        ((df['PEÇA DESCRIÇÃO'].str.contains('_ENG', na=False)) & (df['ESPESSURA'] == '6'))
    )
    
    # Filtrar o DataFrame removendo as linhas que atendem à condição mask_exclude
    df_filtered = df[~mask_exclude]

    # Filtrar e organizar os dados
    relatorio_pecas = df_filtered[['PEÇA DESCRIÇÃO', 'CLIENTE - DADOS DO CLIENTE', 'ALTURA (X)', 'PROF (Y)', 'ESPESSURA', 'AMBIENTE', 'DESENHO']]

    relatorio_pecas = relatorio_pecas.rename(columns={
        'PEÇA DESCRIÇÃO': 'PEÇA DESCRIÇÃO',
        'CLIENTE - DADOS DO CLIENTE': 'CLIENTE',
        'ALTURA (X)': 'ALT (X)',
        'PROF (Y)': 'PROF (Y)',
        'ESPESSURA': 'ESP (Z)',
        'AMBIENTE': 'AMBIENTE',
        'DESENHO': 'DESENHO'
    })
    
    # Adicionar a coluna VISTO
    relatorio_pecas['VISTO'] = ''
    
    # Organizar por ALTURA (X) de forma decrescente e depois por LARGURA (Y) de forma decrescente
    relatorio_pecas = relatorio_pecas.sort_values(by=['ALT (X)', 'PROF (Y)'], ascending=[False, False])

    # Adicionar a coluna NUMERADOR
    relatorio_pecas['NUM'] = range(1, len(relatorio_pecas) + 1)

    # Formatar valores
    data = [relatorio_pecas.columns.tolist()] + formatar_valores(relatorio_pecas.values.tolist())
    
    # Salvar o relatório como PDF
    pasta_arquivo = os.path.dirname(arquivo_excel)  # Obtém o diretório onde o arquivo Excel está localizado
    file_name = os.path.join(pasta_arquivo, 'Relatorio_Pecas.pdf')
    
    # Ajustar margens para usar mais espaço na página
    margins = {'rightMargin': 0.3 * inch, 'leftMargin': 0.3 * inch, 
               'topMargin': 0.3 * inch, 'bottomMargin': 0.3 * inch}
    doc = SimpleDocTemplate(file_name, pagesize=landscape(letter), **margins)
    elements = []

    # Criar a tabela
    table = Table(data)

    # Define o estilo da tabela
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.black),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),   # Justificar à esquerda para 'PEÇA'
        ('ALIGN', (2, 1), (2, -1), 'RIGHT'),  # Justificar à direita para 'ALTURA (X)'
        ('ALIGN', (3, 1), (3, -1), 'RIGHT'),  # Justificar à direita para 'PROF (Y)'
        ('ALIGN', (4, 1), (4, -1), 'RIGHT'),  # Justificar à direita para 'ESPESSURA'
        ('ALIGN', (5, 1), (5, -1), 'LEFT'),   # Justificar à esquerda para 'AMBIENTE'
        ('ALIGN', (6, 1), (6, -1), 'LEFT'),   # Justificar à esquerda para 'DESENHO'
        ('ALIGN', (7, 1), (7, -1), 'RIGHT'),  # Justificar à direita para 'NUMERADOR'
        ('ALIGN', (1, 1), (1, -1), 'LEFT'),   # Justificar à esquerda para 'CLIENTE'
        ('BOX', (0, 0), (-1, -1), 0.2, colors.black),  # Espessura das bordas
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),  # Espessura das linhas de grade
        ('TOPPADDING', (0, 0), (-1, -1), 10),  # Adiciona espaçamento na parte Superior
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),  # Adiciona espaçamento na parte inferior
        ('BOTTOMPADDING', (1, 1), (1, -1), -1),  # Adiciona espaçamento na parte inferior para a coluna 'CLIENTE'
        ('BOTTOMPADDING', (5, 1), (5, -1), -1),  # Adiciona espaçamento na parte inferior para a coluna 'AMBIENTE'
        ('BOTTOMPADDING', (0, 0), (0, -1), 0),  # Adiciona espaçamento na parte inferior para a coluna 'PECA'
    ])

    # Alternar fundo cinza e branco
    num_rows = len(data)
    for i in range(1, num_rows):
        if i % 2 == 1:
            style.add('BACKGROUND', (0, i), (-1, i), colors.Color(0.9, 0.9, 0.9))  # Cinza claro
        else:
            style.add('BACKGROUND', (0, i), (-1, i), colors.white)
    
    table.setStyle(style)

    # Ajustar a largura das colunas para caber no texto
    largura_total = landscape(letter)[0] - (margins['leftMargin'] + margins['rightMargin'])  # Largura total disponível na página
    
    # Definir larguras específicas para as colunas
    largura_colunas = [
        2.0 * inch,  # 'PEÇA'
        2.3 * inch,  # 'CLIENTE'
        0.6 * inch,  # 'ALTURA (X)'
        0.6 * inch,  # 'PROF (Y)'
        0.6 * inch,  # 'ESPESSURA'
        1.9 * inch,  # 'AMBIENTE'
        0.9 * inch,  # 'DESENHO'
        0.8 * inch,  # 'VISTO'
        0.4 * inch,  # 'NUMERADOR'
    ]
    
    # Defina as alturas das linhas
    row_heights = [0.2 * inch] * len(data)  # Define a altura de todas as linhas

    # Configure manualmente a altura das linhas
    for i, height in enumerate(row_heights):
        table._argH[i] = height

    # Ajustar as larguras das colunas para garantir que caibam na largura total disponível
    while sum(largura_colunas) > largura_total:
        for i in range(len(largura_colunas)):
            if largura_colunas[i] > 0.5 * inch:  # Minimizar até um certo ponto
                largura_colunas[i] -= 0.1 * inch
    
    table._argW = largura_colunas

    # Ajustar o tamanho da fonte para que o texto caiba
    pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
    table.setStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
        ('FONTSIZE', (0, 0), (-1, -1), 11),  # Tamanho padrão para todas as colunas
        ('FONTSIZE', (1, 1), (1, -1), 6),    # Tamanho da fonte para a coluna 'CLIENTE'
        ('FONTSIZE', (5, 1), (5, -1), 6),    # Tamanho da fonte para a coluna 'AMBIENTE'
        ('FONTSIZE', (0, 0), (0, -1), 8),    # Tamanho da fonte para a coluna 'PECA'
    ])

    elements.append(table)
    doc.build(elements)

def mostrar_erro(mensagem):
    root = Tk()
    root.withdraw()  # Oculta a janela principal
    messagebox.showerror("Erro", mensagem)
    root.destroy()  # Destroi a janela após mostrar o erro

def main():
    print("  _______    ______   __    __   ______   __    __   ______   ")
    print(" |       \  /      \ |  \  |  \ /      \ |  \  |  \ /      \  ")
    print(" | $$$$$$$\|  $$$$$$\| $$\ | $$|  $$$$$$\| $$\ | $$|  $$$$$$\ ")
    print(" | $$__/ $$| $$__| $$| $$$\| $$| $$__| $$| $$$\| $$| $$__| $$ ")
    print(" | $$    $$| $$    $$| $$$$\ $$| $$    $$| $$$$\ $$| $$    $$ ")
    print(" | $$$$$$$\| $$$$$$$$| $$\$$ $$| $$$$$$$$| $$\$$ $$| $$$$$$$$ ")
    print(" | $$__/ $$| $$  | $$| $$ \$$$$| $$  | $$| $$ \$$$$| $$  | $$ ")
    print(" | $$    $$| $$  | $$| $$  \$$$| $$  | $$| $$  \$$$| $$  | $$ ")
    print("  \$$$$$$$  \$$   \$$ \$$   \$$ \$$   \$$ \$$   \$$ \$$   \$$ ")


    # Abrir a janela de seleção de arquivo
    Tk().withdraw()  # Evitar que a janela principal do Tkinter apareça
    arquivo = askopenfilename(title="Selecione o arquivo Projeto_producao.xls", filetypes=[("Excel files", "*.xls;*.xlsx")])
    
    if arquivo:
        try:
            # Ler o arquivo Excel
            df = pd.read_excel(arquivo)
            
            # Perguntar ao usuário qual relatório deseja gerar
            print("Qual relatório você deseja gerar?")
            print("1. Relatório de Peças")
            opcao = "1"  # input("Digite o número da opção: ")
            if opcao == '1':
                gerar_relatorio_pecas(df, arquivo)
            else:
                print("Opção inválida.")
        except Exception as e:
            mostrar_erro(f"Erro ao ler o arquivo ou gerar o relatório: {e}")

if __name__ == "__main__":
    main()
