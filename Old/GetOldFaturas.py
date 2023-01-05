# Extract keywords from multiple PDF files, create a dataframe, then export it to an .xlsx file.

from concurrent.futures import process
import os  # provides functions for creating and removing a directory (folder), fetching its contents, changing and identifying the current directory
import pandas as pd  # flexible open source data analysis/manipulation tool
import glob  # generates lists of files matching given patterns
import pdfplumber  # extracts information from .pdf documents
import os
import pathlib
from tkinter import filedialog
from tkinter import *
import re
"""
Obtain key words from repetitive documents, then extract as a dataframe to an .xlsx !
"""

# defining the functions used in main()
def get_keyword(start, end, text):
    """
    start: should be the word prior to the keyword.
    end: should be the word that comes after the keyword.
    text: represents the text from the page(s) you've just extracted.
    """
    for i in range(len(start)):
        try:
            field = (text.split(start[i]))[1].split(end[i])[0]
            return field
        except:
            continue
        ##REFORMULAR A PESQUISA
def testKeyword(word):
    if (word is None) or (len(word)==0):
        return ""
    else:
        newWord = word.split(None,7)
        return newWord[5]

        ##Confirma se incidiu ICMS sobre o TUSD, N = NUMERO DO ITEM DA LISTA A SER DEVOLVIDO 8 = ICMS PAGO, 7 = ALICOTA ICMS
def isTax(a,n):
        if(a is None) or (len(a)==0):
            return ""
        else:
            m = re.search(r'\n', a)
            b = m.span()
            c = a[:int(b[1])].split()
            if(len(c) > 9):
                return c[n]
            else:
                return ""


def main():
    #Uma interface para pedir para o usuário selecionar o caminho das pastas que estão o pdf
    root = Tk()
    root.withdraw()
    #Salva o caminho da pasta 
    currentPath = str(filedialog.askdirectory())
    #Adiciona a função *.pdf para abrir todas as pastas do diretório
    pdfPath = currentPath+"\*.pdf"

    # create an empty dataframe, from which keywords from multiple .pdf files will be later appended by rows.
    my_dataframe = pd.DataFrame()
    for files in glob.glob(pdfPath):
        with pdfplumber.open(files) as pdf:
            #getting the name of the file
            PDFname = str(os.path.basename(files))
            page = pdf.pages[0]
            text = page.extract_text()

            ##text = " ".join(text.split())
            n = 0
            my_list =[]
            # use the function get_keyword as many times to get all the desired keywords from a pdf document.
            
            # Separa entre Empresa e Residencial
            start = ["www.cpflempresas.com.br"]
            end = ["Série"]
            isResidencial = get_keyword(start, end, text)
            if((isResidencial is None) or (len(isResidencial)==0)):
                # obtain Número da Fatura #1
                start = ["Nº"]
                end = ["Série"]
                nFatura = get_keyword(start, end, text)
                if(((nFatura is None) or (len(nFatura)==0))):
                    nFatura = PDFname
                else:
                    #obtain  ALICOTA ICMS
                    start = ["Disp Sistema-TE"]
                    end = ["Subtotal"]
                    ALicms = isTax(get_keyword(start, end, text),6)
                    if(ALicms == ""):
                        start = ["Cons Ponta - TE"]
                        end = ["Subtotal"]
                        ALicms = isTax(get_keyword(start, end, text),6)
                        if(ALicms == ""):
                            start = ["Consumo - TE"]
                            end = ["Subtotal"]
                            ALicms = isTax(get_keyword(start, end, text),6)
                    
                    n = 1

                    # obtain Aliquota ICMS #3
                    start = ["0800 010 2570"]
                    end = ["Descrição"]
                    tempHolder = get_keyword(start, end, text)
                    #Se não encontrar o primeiro telefone, procura pelo próximo
                    if((tempHolder is None) or (len(tempHolder)==0)):
                        start = ["0800 010 1010"]
                        end = ["Descrição"]
                        tempHolder = get_keyword(start, end, text)
                        y = tempHolder.split(None, 7)
                        mes = y[2]
                        ident = y[6]
                        del y
                    else:
                        #extraindo identificação e mês
                        y = tempHolder.split(None, 7)
                        mes = y[2]
                        ident = y[6]
                        #Procurando pelo ICMS através do mês e abreviação do ano
                        del y
                    

            #Especial para faturas de empresas
            else:
                # obtain Número da Fatura #1
                start = ["Nº."]
                end = ["série"]
                nFatura = get_keyword(start, end, text)

                ##obtain  ALICOTA ICMS
                start = ["Disp Sistema-TE"]
                end = ["Subtotal"]
                ALicms = isTax(get_keyword(start, end, text),7)
                if(ALicms == ""):
                    start = ["Cons Ponta - TE"]
                    end = ["Subtotal"]
                    ALicms = isTax(get_keyword(start, end, text),7)
                    if(ALicms == ""):
                        start = ["Consumo - TE"]
                        end = ["Subtotal"]
                        ALicms = isTax(get_keyword(start, end, text),7)

                # obtain MÊS E IDENTIFICAÇÃO
                start = ["0800 770 4140"]
                end = ["Descrição"]
                tempHolder = get_keyword(start, end, text)
                y = tempHolder.split(None, 5)
                mes = y[2]
                ident = y[1]
                del tempHolder
                del y
                """tempList = list(mes.split("/"))
                tempList1 = tempList[1][2:]
                y = tempList[0]+'/'+tempList1
                del tempHolder
                del tempList1
                del tempList
                start = [y]
                end = ["Total"]
                tempHolder = get_keyword(start, end, text)
                z = tempHolder.split(None, 8)
                ALicms = z[6]
                del tempHolder
                del z
                del y"""

            # create a list with the keywords extracted from current document.
            if(nFatura == PDFname):
                for i in range(10):
                    my_list.append("")
                my_list.insert(2, nFatura)
                processFailed = True
            else:
                # obter ICMS PAGO e Cálculo #2
                start = ["Total Consolidado"]
                end = ["Consumo"]
                tempHolder = get_keyword(start, end, text)
                #Caso o valor não seja encontrado, será procurado em outra página
                if((tempHolder is None)) or (len(tempHolder)==0):
                    page1 = pdf.pages[1]
                    text1 = page1.extract_text()
                    text1 = " ".join(text1.split())
                    tempHolder = get_keyword(start, end, text1)
                    y = tempHolder.split(None, 4) 
                    BCicms = y[1]
                    ICMSpago = y[2]
                    del tempHolder
                    del y
                else:
                    y = tempHolder.split(None, 4)
                    BCicms = y[1]
                    ICMSpago = y[2]
                    del tempHolder
                    del y

                # obtain Consumo Ponta TUSD #4
                start = ["Consumo Ponta [KWh] - TUSD"]
                end = ["Subtotal"]
                CPtusd = isTax(get_keyword(start, end, text),(8-n))

                    
                # obtain Consumo Fora Ponta TUSD #5
                start = ["Consumo Uso Sistema [KWh]-TUSD"]
                end = ["Subtotal"]
                CFPtusd = isTax(get_keyword(start, end, text),(8-n))
                if(CFPtusd == ""):
                    start = ["Consumo Fora Ponta [KWh]-TUSD"]
                    end = ["Subtotal"]
                    CFPtusd = isTax(get_keyword(start, end, text),(8-n))
                    if(CFPtusd ==""):
                        start = ["Custo Disp Uso Sistema TUSD"]
                        end = ["Subtotal"]
                        CFPtusd = isTax(get_keyword(start, end, text),(8-n))
                
                # obtain Demanda TUSD #6
                start = ["Demanda [kW] - TUSD"]
                end = ["Subtotal"]
                Dtusd = isTax(get_keyword(start, end, text),(8-n))
                
                # obtain Demanda Ponta TUSD  #7
                start = ["Demanda Ponta [kW] - TUSD"]
                end = ["Subtotal"]
                DPtusd = isTax(get_keyword(start, end, text),(8-n))
                

                # obtain Demanda  Fora Ponta #8
                start = ["Demanda F Ponta [kW] -TUSD"]
                end = ["Subtotal"]
                DFPtusd = isTax(get_keyword(start, end, text),(8-n))


                my_list = [mes, ident, nFatura, BCicms, ALicms, ICMSpago, CPtusd, CFPtusd, Dtusd, DPtusd, DFPtusd]
                processFailed = False
            # append my list as a row in the dataframe.
            my_list = pd.Series(my_list)

            # append the list of keywords as a row to my dataframe.
            my_dataframe = my_dataframe.append(my_list, ignore_index=True)
            if(processFailed):
                print("Não foi possível extrair os dados do arquivo '"+PDFname+"'!")
            else:
                print("Os dados do arquivo '"+PDFname+"' foram extraídos com sucesso!")

    # rename dataframe columns using dictionaries.
    my_dataframe = my_dataframe.rename(
        columns={
            0: "Mês",
            1: "Identificador",
            2: "N° Fatura",
            3: "Base de Cálculo ICMS Pago",
            4: "Alíquota ICMS",
            5: "ICMS Pago",
            6: "Consumo Ponta [KWh] - TUSD",
            7: "Consumo Fora Ponta [KWh]-TUSD",
            8: "Demanda [kW] - TUSD",
            9: "Demanda Ponta [kW] - TUSD",
            10: "Demanda F Ponta [kW] - TUSD",
        }
    )

    # change my current working directory
    os.chdir(currentPath)

    # extract my dataframe to an .xlsx file!
    my_dataframe.to_excel("sample_excel.xlsx", sheet_name="my dataframe")
    print("")
    print(my_dataframe)


if __name__ == "__main__":
    main()
