import os
import sys
import datetime
from PIL import Image, ImageDraw, ImageFont
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd


def get_template_path(filename):
    """
    Retorna o caminho dos templates, para o caso do executável gerado na dist.
    """
    if hasattr(sys, '_MEIPASS'):
        # Executando a partir do executável do PyInstaller na dist
        base_path = sys._MEIPASS
    else:
        # Executando a partir do código-fonte (desenvolvimento)
        base_path = os.path.abspath(".")

    return os.path.join(base_path, 'imagens', filename)


def obter_data_outorga():
    """
    Retorna a data inserida pelo usuário ou, se vazia, a data atual no formato 'Brasília, DD de Mês de YYYY'.
    
    Returns:
        str: Data formatada.
    """
    data = data_outorga.get()
    
    if not data:
        data_atual = datetime.datetime.now()

        meses = {
            1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
            7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
        }

        data_formatada = f"Brasilia, {data_atual.day} de {meses[data_atual.month]} de {data_atual.year}"
        return data_formatada
    
    return data


def desenha_texto_centralizado(p1_coord: tuple, p2_coord: tuple, fonte: ImageFont, texto: str, draw: ImageDraw):
    """
    Desenha texto centralizado, de acordo com coordenadas delimitadas por um retângulo arbitrário.

    Args:
        p1_coord (tuple): Coordenadas da ponta superior esquerda do retângulo.
        p2_coord (tuple): Coordenadas da ponta inferior direita do retângulo.
        fonte (ImageFont): Objeto fonte do texto.
        texto (str): Texto a ser exibido. 
        draw (ImageDraw): Objeto Draw que modifica o objeto Image, que representa o arquivo .jpg.
    """
    # Obtém largura e altura do texto dentro do retângulo delimitador
    bbox = draw.textbbox((0, 0), texto, font=fonte)
    largura_texto, altura_texto = bbox[2] - bbox[0], bbox[3] - bbox[1]

    # Calcula coordenadas do texto centralizado
    x = (p1_coord[0] + p2_coord[0] - largura_texto) // 2
    y = (p1_coord[1] + p2_coord[1] - altura_texto) // 2

    # Desenha o texto centralizado
    draw.text((x, y), texto, fill='black', font=fonte)


def switch_template(path_to_template:str):
    """
    Alterna emissor da outorga para Rafaello.

    Args:
        path_to_template (str): Caminho do template a ser mascarado.
    Returns:
        image: Objeto Image com emissor alternado.
    """
    # Gera objeto draw
    image = Image.open(path_to_template)
    draw = ImageDraw.Draw(image)

    # Mascara nome do emissor
    draw.rectangle([523, 649, 778, 687], fill='white')

    # Substitui nome do emissor para Rafaello
    fonte_times = ImageFont.truetype('Timesbd.ttf', 12)
    draw.text((520, 650), 'Rafaello Pinheiro Mazzoccante - Faixa Preta 2° Dan', fill='black', font=fonte_times)
    fonte_times = ImageFont.truetype('times.ttf', 10)
    draw.text((565, 664), 'Cadastro Nacional Judô - AT039278/CBJ', fill='black', font=fonte_times)
    draw.text((605, 674), 'CREF 009236-G/DF', fill='black', font=fonte_times)

    return image


def gerar_jpg(font_size: str, cor: str, nome_outorga: str, data_outorga: str, emissor: str, manual=True):
    """
    Gera diploma em formato .jpg na pasta 'imagens_geradas'.

    Args:
        font_size (str): Tamanho da fonte.
        cor (str): Cor da faixa selecionada.
        nome_outorga (str): Nome do diplomado.
        data_outorga (str): Data de outorga.
        emissor (str): Nome do emissor.
        manual (bool): Indica se a geração foi manual (True) ou via importação de Excel (False).
    """

    # Verifica a faixa selecionada e usa o template correspondente
    if cor == 'Branca (Ponta Cinza)':
        path_to_file = get_template_path('template_branca_pcinza.jpg')
    elif cor == 'Cinza':
        path_to_file = get_template_path('template_cinza.jpg')
    elif cor == 'Cinza (Ponta Azul)':
        path_to_file = get_template_path('template_cinza_pazul.jpg')
    elif cor == 'Azul':
        path_to_file = get_template_path('template_azul.jpg')
    elif cor == 'Azul (Ponta Amarela)':
        path_to_file = get_template_path('template_azul_pamarela.jpg')
    elif cor == 'Amarela':
        path_to_file = get_template_path('template_amarela.jpg')
    elif cor == 'Amarela (Ponta Laranja)':
        path_to_file = get_template_path('template_amarela_plaranja.jpg')
    elif cor == 'Laranja':
        path_to_file = get_template_path('template_laranja.jpg')
    elif cor == 'Verde':
        path_to_file = get_template_path('template_verde.jpg')
    elif cor == 'Roxa':
        path_to_file = get_template_path('template_roxa.jpg')
    elif cor == 'Marrom':
        path_to_file = get_template_path('template_marrom.jpg')
    else:
        messagebox.showerror("Erro", "Cor da faixa inválida.")
        return

    # Abre o template
    image = Image.open(path_to_file)

    # Gera o objeto Draw
    draw = ImageDraw.Draw(image)

    # Desenha a data de emissão
    fonte = ImageFont.truetype('Tokyo.ttf', 20)
    desenha_texto_centralizado((300, 540), (640, 572), fonte, data_outorga, draw)

    # Desenha o nome do diplomado
    fonte = ImageFont.truetype('Tokyo.ttf', int(font_size))
    desenha_texto_centralizado((500, 260), (1010, 290), fonte, nome_outorga, draw)

    # Salva a imagem na pasta 'imagens_geradas'
    output_path = os.path.join('imagens_geradas', f'{nome_outorga.upper()}.jpg')
    image.save(output_path)

    # Se o emissor for Rafaello, aplica a máscara sobre o nome do emissor
    if emissor == 'Rafaello':
        switch_template(output_path).save(output_path)

    # Notifica o usuário em caso de geração manual
    if manual:
        messagebox.showinfo("Sucesso", "Diplomas gerados com sucesso!")


def retornar_selecoes() -> tuple:
    """
    Retorna valores inseridos pelo usuário.

    Returns:
        tuple: Tupla com todas entradas do usuário.
    """
    # Lida com seleção da Check List
    if (braulio_c1_variable.get() == 1):
        emissor = 'Bráulio'
    elif (rafaello_c1_variable.get() == 1):
        emissor = 'Rafaello'

    return (combo_fontes.get(), combo_cores.get(), 
            nome_outorga.get(), data_outorga.get(), 
            emissor)


def gerar_jpg_retransmitter():
    """
    Supera limitação de evocação de eventos do Tkinter, em que não é possível definir diretamente argumentos
    da função associada a um widget. Obtém entradas e chama função que gera diploma.
    """
    # Obtém entradas de seleção
    font_size, cor, nome, data, emissor = retornar_selecoes()

    # Obtém a data do sys
    data = obter_data_outorga()

    # Gera diploma jpg, armazenado em 'imagens_geradas'
    gerar_jpg(font_size, cor, nome, data, emissor)


def desabilitar_checkboxes():
    if (braulio_c1_variable.get()) == 1:
        rafaello_c1.deselect()
    elif (rafaello_c1_variable.get()) == 1:
        braulio_c1.deselect()


def importar_excel():
    """
    Gera diplomas ao importar informações contidas em uma planilha do Excel, 
    usando a data especificada na label de data (se disponível) ou a data atual.
    """
    file_path = filedialog.askopenfilename(title="Selecione o arquivo Excel",
                                           filetypes=(("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")))

    if file_path:
        try:
            df = pd.read_excel(file_path)

            # Verifica as colunas "Nome do Aluno" e "Faixa"
            if "Nome do Aluno" not in df.columns or "Faixa" not in df.columns:
                messagebox.showerror("Erro", "A planilha deve conter as colunas 'Nome do Aluno' e 'Faixa'.")
                return

            # Decide emissor da outorga
            if braulio_c1_variable.get() == 1:
                emissor = "Bráulio"
            elif rafaello_c1_variable.get() == 1:
                emissor = "Rafaello"
            else:
                messagebox.showerror("Erro", "Selecione um assinante (Bráulio ou Rafaello).")
                return

            # Obtém a data de outorga (usa a função que verifica a label)
            data = obter_data_outorga()

            for _, row in df.iterrows():
                nome = row["Nome do Aluno"]
                faixa = row["Faixa"]

                # Gera os diplomas usando a data especificada (ou gerada automaticamente)
                gerar_jpg(combo_fontes.get(), faixa, nome, data, emissor, False)
                
            messagebox.showinfo("Sucesso", "Diplomas gerados com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o arquivo Excel: {e}")


if __name__ == "__main__":
    # Setup Inicial Interface Gráfica

    # Janela Master
    window = tk.Tk()
    window.title('Editor de Outorgas de Faixa')
    window.geometry("525x165")
    window.resizable(False, False)
    window.configure(bg='gray')

    texto_nome_outorga = tk.Label(text='Insira o nome:', bg='gray')
    texto_nome_outorga.place(x=7, y=13)

    # Nome Outorga
    nome_outorga = tk.StringVar()
    textArea_outorga = tk.Entry(master=window, width=30,
                                font='Calibri 15',
                                textvariable=nome_outorga)
    textArea_outorga.place(x=100, y=9)

    texto_data_outorga = tk.Label(text='Insira a data:', bg='gray')
    texto_data_outorga.place(x=7, y=53)

    # Data Outorga
    data_outorga = tk.StringVar()
    textArea_data = tk.Entry(master=window, width=41,
                            font='Calibri 15',
                            textvariable=data_outorga)
    textArea_data.place(x=100, y=49)

    # Botao Gerar.jpg
    botao = tk.Button(master=window, 
                    width=36,
                    height=1,
                    text='Gerar .jpg',
                    font='Calibri 9 bold',
                    command=gerar_jpg_retransmitter)
    botao.place(x=290, y=93)

    # Botao importar excel
    botao_importar = tk.Button(master=window, width=83, height=1, text='Importar Excel', font='Calibri 9 bold', command=importar_excel)
    botao_importar.place(x=7, y=130)

    # CheckBox - Bráulio
    braulio_c1_variable = tk.IntVar()
    braulio_c1 = tk.Checkbutton(master=window, 
                                text='Bráulio', 
                                variable=braulio_c1_variable, 
                                onvalue=1,
                                bg='gray',
                                command=desabilitar_checkboxes)
    braulio_c1.place(x=145, y=92)

    # CheckBox - Rafaello
    rafaello_c1_variable = tk.IntVar()
    rafaello_c1 = tk.Checkbutton(master=window, 
                                text='Porto', 
                                variable=rafaello_c1_variable,
                                onvalue=1,
                                bg='gray',
                                command=desabilitar_checkboxes)
    rafaello_c1.place(x=210, y=92)

    # ComboBox - Font Size
    font_sizes = [25, 30, 35]
    combo_fontes = ttk.Combobox(master=window, values=font_sizes, width=12)
    combo_fontes.place(x=420, y=13)

    # ComboBox - Cor Da Faixa
    cores = ['Branca (Ponta Cinza)', 'Cinza', 'Cinza (Ponta Azul)', 'Azul', 
            'Azul (Ponta Amarela)', 'Amarela', 'Amarela (Ponta Laranja)',
            'Laranja', 'Verde', 'Roxa', 'Marrom']
    combo_cores = ttk.Combobox(master=window, values=cores, width=17)
    combo_cores.place(x=7, y=94)

    window.mainloop()
