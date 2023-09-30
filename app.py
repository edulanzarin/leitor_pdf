import PyPDF2
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from ttkthemes import ThemedStyle
import threading

root = tk.Tk()
root.title("Conversor de Relatórios")
root.geometry("600x300")

icon_path = r"C:\Users\edula\OneDrive\Documentos\Eduardo\IDEAL\Lanzarin\code.ico"

root.iconbitmap(icon_path)


def processar_pdf_thread():
    # Desabilite o botão de processamento enquanto o PDF está sendo processado
    process_button.config(state="disabled")
    progress_bar.pack(pady=20)

    with open(file_path, "rb") as pdf_file:
        dados_pdf = PyPDF2.PdfReader(pdf_file)

        pdf_file = None
        registros_capital_emporio = []
        registros_qualitplacas = []
        registros_lojao = []

        global selected_empresa_var
        selected_empresa = selected_empresa_var.get()

        quebra_condicao = ""

        if selected_empresa == "Empório Astral":
            quebra_condicao = "04-EMP"

        elif selected_empresa == "Capital Six":
            quebra_condicao = "CAPITAL SIX"

        elif selected_empresa == "Qualitplacas":
            quebra_condicao = "QUALITPLACAS"

        elif selected_empresa == "Lojão Astral":
            quebra_condicao = "09-LOJÃO"
        # Defina df como uma lista vazia antes dos blocos if/elif
        df = pd.DataFrame()

        linhas_imprimir = False

        if quebra_condicao == "CAPITAL SIX" or quebra_condicao == "04-EMP":
            total_pages = len(dados_pdf.pages)
            current_page = 0

            for pagina in dados_pdf.pages:
                texto_pagina = pagina.extract_text()
                linhas = texto_pagina.split("\n")
                # É o Padrão 2, processar 6 linhas após a quebra
                for linha in linhas[5:]:
                    # linha = linha.replace("PAGAMENTO REALIZADO", "")

                    partes = linha.split()
                    if (
                        len(partes) >= 5
                        and "A VISTA" not in linha
                        and "BOLETO 1" not in linha
                        and "DESPESAS FIXAS" not in linha
                        and "DESPESAS VARIAVEIS" not in linha
                        and "ESCRITÓRIO DESPESAS" not in linha
                        and "INVESTIMENTO" not in linha
                        and "RETIRADA SOCIOS" not in linha
                        and "SEM PORTADOR" not in linha
                        and "CONSÓRCIOS" not in linha
                        and "FEIRAS" not in linha
                    ):
                        numero_duplicata = partes[0]
                        nome_empresa = " ".join(
                            partes[1:-4] + ["NF" + numero_duplicata]
                        )
                        data_vencimento = "N/A"
                        valor = partes[-1]

                        valor = (
                            valor.replace(".", "")
                            .replace(",", ".")
                            .replace(".00", "")
                            .replace(".0", "")
                        )
                        substituicoes = [
                            ".10",
                            ".20",
                            ".30",
                            ".40",
                            ".50",
                            ".60",
                            ".70",
                            ".80",
                            ".90",
                        ]  # Adicione mais substituições se necessário
                        for substituicao in substituicoes:
                            if valor.endswith(substituicao):
                                valor = valor[:-2] + substituicao[-2]

                        current_page += 1
                        progress_value = (current_page / total_pages) * 100
                        progress_bar["value"] = progress_value
                        root.update_idletasks()

                        registros_capital_emporio.append(
                            {
                                "DATA": data_vencimento,
                                "FORNECEDOR": nome_empresa,
                                "VALOR": valor,
                            }
                        )

            df = pd.DataFrame(registros_capital_emporio)
            progress_bar["value"] = 100

        elif quebra_condicao == "QUALITPLACAS":
            linhas_impressas = 0
            linha_anterior_valor_zero = False

            valores_por_dia = {}  # Dicionário para armazenar os valores por dia

            total_pages = len(dados_pdf.pages)
            current_page = 0

            for pagina_num, pagina in enumerate(dados_pdf.pages, 1):
                texto_pagina = pagina.extract_text()
                linhas = texto_pagina.split("\n")

                quebra_pagina = False
                linhas_quebra = 0

                for linha in linhas:
                    if quebra_condicao in linha:
                        quebra_pagina = True
                        linhas_quebra = 0

                    if (
                        quebra_pagina and pagina_num == 1
                    ):  # Apenas para a primeira página
                        linhas_quebra += 1
                        if (
                            linhas_quebra >= 8
                            and "Histórico" not in linha
                            and "Complemento" not in linha
                        ):
                            linhas_quebra = 0
                            linhas_imprimir = True
                            quebra_pagina = False

                    if linhas_imprimir:
                        if (
                            "Histórico" not in linha
                            and "Complemento" not in linha
                            and "Página:" not in linha
                        ):
                            partes = linha.split()
                            if len(partes) > 2:
                                if (
                                    "BANCO SAFRA MATRIZ" in linha
                                    or "9VIACREDI" in linha
                                ):
                                    partes = [partes[0]] + partes[-3:] + partes
                                    data = partes[0]
                                    if (
                                        "RECEITA DE REBATE" not in linha
                                        and "DEVOLUÇÃO DE COMPRA" not in linha
                                    ):
                                        valor = partes[2]
                                        credito = ""
                                    else:
                                        valor = partes[1]
                                        credito = partes[1]

                                    valor = (
                                        valor.replace(".", "")
                                        .replace(",", ".")
                                        .replace(".00", "")
                                        .replace(".0", "")
                                    )
                                    substituicoes = [
                                        ".10",
                                        ".20",
                                        ".30",
                                        ".40",
                                        ".50",
                                        ".60",
                                        ".70",
                                        ".80",
                                        ".90",
                                    ]  # Adicione mais substituições se necessário
                                    for substituicao in substituicoes:
                                        if valor.endswith(substituicao):
                                            valor = valor[:-2] + substituicao[-2]
                                    credito = (
                                        credito.replace(".", "")
                                        .replace(",", ".")
                                        .replace(".00", "")
                                        .replace(".0", "")
                                    )
                                    substituicoes = [
                                        ".10",
                                        ".20",
                                        ".30",
                                        ".40",
                                        ".50",
                                        ".60",
                                        ".70",
                                        ".80",
                                        ".90",
                                    ]  # Adicione mais substituições se necessário
                                    for substituicao in substituicoes:
                                        if credito.endswith(substituicao):
                                            credito = credito[:-2] + substituicao[-2]
                                    if (
                                        valor != "0"
                                        and "DESPESAS BANCARIA" not in linha
                                        and "ALUGUEIS MAQUINA" not in linha
                                    ):
                                        if valor == credito:
                                            linha_anterior_valor_zero = False
                                            if data in valores_por_dia:
                                                valores_por_dia[data] -= float(valor)
                                            valor = ""

                                        else:
                                            linha_anterior_valor_zero = False
                                            if data not in valores_por_dia:
                                                valores_por_dia[data] = float(valor)
                                            else:
                                                valores_por_dia[data] += float(valor)
                                    else:
                                        linha_anterior_valor_zero = True
                                else:
                                    if linha_anterior_valor_zero:
                                        linha_anterior_valor_zero = False
                                    else:
                                        if "REC.REF.DOC.:" in linha:
                                            partes[0] = partes[0].replace(
                                                "REC.REF.DOC.:", ""
                                            )
                                        if "PAG.REF.DOC.:" in linha:
                                            partes[0] = partes[0].replace(
                                                "PAG.REF.DOC.:", ""
                                            )
                                        if (
                                            "PAG.REF.DOC.:AGR" not in linha
                                            and "REC.REF.DOC.:AGR" not in linha
                                        ):
                                            if "-" in partes[0]:
                                                partes[0] = (
                                                    partes[0].split("-", 1)[0].lstrip()
                                                )
                                        else:
                                            partes[0] = partes[0].replace(
                                                "PAG.REF.DOC.:AGR", ""
                                            )
                                            partes[0] = partes[0].replace(
                                                "REC.REF.DOC.:AGR", ""
                                            )
                                            if "-" in partes[0]:
                                                partes[0] = (
                                                    partes[0].split("-", 1)[1].lstrip()
                                                )
                                        if "SACADO" in linha:
                                            partes[1] = partes[1].replace("SACADO:", "")
                                        if "DESC.TITULO" in linha:
                                            partes[0] = partes[0].replace(
                                                "DESC.TITULO", ""
                                            )
                                            if "-" in partes[1]:
                                                partes[1] = (
                                                    partes[1].split("-", 1)[0].lstrip()
                                                )
                                            if "/" in partes[1]:
                                                partes[1] = (
                                                    partes[1].split("/", 1)[0].lstrip()
                                                )
                                            nota = partes[1]
                                            fornecedor = "DESCONTO DO TITULO " + nota
                                        else:
                                            if "-" in partes[1]:
                                                partes[1] = (
                                                    partes[1].split("-", 1)[1].lstrip()
                                                )

                                        if "PAG.REF.DOC.: " in linha:
                                            fornecedor = " ".join(partes[2::])
                                            nota = partes[1]
                                        else:
                                            fornecedor = " ".join(partes[1:])
                                            nota = partes[0]

                                            current_page += 1
                                            progress_value = (
                                                current_page / total_pages
                                            ) * 100
                                            progress_bar["value"] = progress_value
                                            root.update_idletasks()

                                        registros_qualitplacas.append(
                                            {
                                                "DATA": data,
                                                "FORNECEDOR": fornecedor + " " + nota,
                                                "VALOR": valor,
                                                "CREDITO": credito,
                                            }
                                        )

                            linhas_impressas += 1

            df = pd.DataFrame(registros_qualitplacas)
            progress_bar["value"] = 100

        elif quebra_condicao == "09-LOJÃO":
            pagamentos_fornecedor_atual = []
            quebra_pagina = False
            linhas_quebra = 0
            contador_linhas = 0
            fornecedor_atual = None

            total_pages = len(dados_pdf.pages)
            current_page = 0

            for pagina_num, pagina in enumerate(dados_pdf.pages, 1):
                texto_pagina = pagina.extract_text()
                linhas = texto_pagina.split("\n")
                for linha in linhas:
                    if quebra_condicao in linha:
                        quebra_pagina = True
                        linhas_quebra = 0
                        contador_linhas = 0

                    if quebra_pagina:
                        linhas_quebra += 1
                        if linhas_quebra >= 4:
                            linhas_imprimir = True
                            quebra_pagina = False

                    if (
                        linhas_imprimir
                        and contador_linhas >= 4
                        and "Total da pessoa" not in linha
                    ):
                        if "Pessoa:" in linha:
                            if fornecedor_atual and pagamentos_fornecedor_atual:
                                registros_lojao.extend(pagamentos_fornecedor_atual)
                            match = re.search(r"Pessoa:\s+(\d+\s+)?(.+)", linha)
                            if match:
                                fornecedor_nome = match.group(2).strip()
                                fornecedor_atual = fornecedor_nome
                                pagamentos_fornecedor_atual = []
                        elif fornecedor_atual:
                            partes = linha.split()
                            if len(partes) >= 11:
                                datas = re.findall(r"\d{2}/\d{2}/\d{4}", linha)
                                valores = re.findall(
                                    r"\d+(?:\.\d+,\d+|\.\d+,\d+|\,\d+)", linha
                                )
                                # Realizar substituições em uma lista de valores
                                valores = [
                                    valor.replace(".", "")
                                    .replace(",", ".")
                                    .replace(".00", "")
                                    .replace(".0", "")
                                    for valor in valores
                                ]
                                substituicoes = [
                                    ".00",
                                    ".10",
                                    ".20",
                                    ".30",
                                    ".40",
                                    ".50",
                                    ".60",
                                    ".70",
                                    ".80",
                                    ".90",
                                ]  # Adicione mais substituições se necessário
                                if valores and len(valores) > 1:
                                    for i in range(len(valores)):
                                        valor = valores[i]
                                        for substituicao in substituicoes:
                                            if valor.endswith(substituicao):
                                                valores[i] = (
                                                    valor[:-2] + substituicao[-2]
                                                )
                                else:
                                    valores = [""] * len(
                                        valores
                                    )  # Crie uma lista vazia com o mesmo comprimento

                                # Os valores resultantes estarão na lista 'valores'

                                current_page += 1
                                progress_value = (current_page / total_pages) * 100
                                progress_bar["value"] = progress_value
                                root.update_idletasks()

                                pagamentos_fornecedor_atual.append(
                                    {
                                        "DATA": datas[2]
                                        if datas and len(datas) > 2
                                        else "",
                                        "FORNECEDOR": fornecedor_atual
                                        + " "
                                        + "NF"
                                        + partes[1],
                                        "VALOR": (valores[1])
                                        if valores and len(valores) > 1
                                        else "",
                                    }
                                )

                    contador_linhas += 1

                if fornecedor_atual and pagamentos_fornecedor_atual:
                    registros_lojao.extend(pagamentos_fornecedor_atual)

            df = pd.DataFrame(registros_lojao)
            progress_bar["value"] = 100

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
    )
    if not save_path:
        status_label.config(text="Nenhum local de salvamento selecionado.")
        return

    df.to_excel(save_path, index=False)

    messagebox.showinfo("Concluído", f"O arquivo excel salvo em:\n{save_path}")

    root.after(10, finish_processing)
    root.quit()


def processar_pdf():
    if not file_path:
        status_label.config(text="Nenhum arquivo PDF selecionado.")
        return

    pdf_thread = threading.Thread(target=processar_pdf_thread)
    pdf_thread.start()


def selecionar_pdf():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        file_name = os.path.basename(file_path)
        status_label.config(text=f"{file_name}")
    else:
        status_label.config(text="Nenhum arquivo PDF selecionado.")


def change_cursor(event):
    select_button.config(cursor="hand2")
    process_button.config(cursor="hand2")
    empresa_menu.config(cursor="hand2")


def selecionar_empresa():
    global selected_empresa_var
    empresas = ["Qualitplacas", "Lojão Astral", "Capital Six", "Empório Astral"]

    selected_empresa_var = tk.StringVar(root)
    selected_empresa_var.set(empresas[0])

    global empresa_menu
    empresa_menu = ttk.Combobox(
        root, textvariable=selected_empresa_var, values=empresas
    )
    empresa_menu.bind("<Enter>", change_cursor)
    empresa_menu.pack(pady=20)


selecionar_empresa()

select_button = ttk.Button(
    root, text="Selecionar PDF", command=selecionar_pdf, width=20, padding=10
)
select_button.bind("<Enter>", change_cursor)

select_button.pack(pady=20)

# quebra_condicao_label = tk.Label(root, text="Condição para quebrar a página:")
# quebra_condicao_label.pack()

# quebra_condicao_entry = tk.Entry(root, width=40)
# quebra_condicao_entry.pack()

process_button = ttk.Button(
    root, text="Processar PDF", command=processar_pdf, width=20, padding=10
)
process_button.bind("<Enter>", change_cursor)
process_button.pack(pady=20)


def finish_processing():
    status_label.config(text="Processamento concluído.")
    process_button.config(state="normal")


progress_bar = ttk.Progressbar(root, mode="determinate")
progress_bar.pack_forget()

status_label = tk.Label(root, text="")
status_label.pack()

style = ThemedStyle(root)
style.set_theme("vista")

root.mainloop()
