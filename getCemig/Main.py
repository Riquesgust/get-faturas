# Extract keywords from multiple PDF files, create a dataframe, then export it to an .xlsx file.

from concurrent.futures import process
from pathlib import Path
from xml.etree.ElementInclude import include  # provides functions for creating and removing a directory (folder), fetching its contents, changing and identifying the current directory
import pandas as pd  # flexible open source data analysis/manipulation tool
import glob  # generates lists of files matching given patterns
import pdfplumber  # extracts information from .pdf documents
import os
import pathlib
from tkinter import filedialog
from tkinter import *


"""
Obtain key words from repetitive documents, then extract as a dataframe to an .xlsx !
"""
"""
Esse programa abre os PDF e extrai as informações e retorna uma planilha no padrão da comissão
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
        
def main():
    # create an empty dataframe, from which keywords from multiple .pdf files will be later appended by rows.
    my_dataframe = pd.DataFrame()

    root = Tk()
    root.withdraw()
    #Salva o caminho da pasta 
    currentPath = str(filedialog.askdirectory())
    #Adiciona a função *.pdf para abrir todas as pastas do diretório
    pdfPath = currentPath+"\*.pdf"

    for files in glob.glob(pdfPath):
        with pdfplumber.open(files) as pdf:
            #getting the name of the file
            PDFname = str(os.path.basename(files))
            page = pdf.pages[0]
            text = page.extract_text()
            
            text = " ".join(text.split())
            my_list =[]
            # use the function get_keyword as many times to get all the desired keywords from a pdf document.
            
            # Separa entre DTCEA-TRM e Globalizada
            start = ["INSTALAÇÃO"]
            end = ["Classe"]
            isNotGlobal = get_keyword(start, end, text)
            
            # Caso seja fatura Globalizada executa esse trecho
            if((isNotGlobal is None) or (len(isNotGlobal)==0)):
                
                # obtain ident, mes, vencimento, valorLiq #1
                start = ["CONTA CONTRATO"]
                end = ["DOCUMENTO"]
                cemigHeader = get_keyword(start, end, text)
                cemigHeader = cemigHeader.split(None)
                ident = cemigHeader[8]
                mes = cemigHeader[23]
                dataVenc = cemigHeader[24]
                valorLiq = float(cemigHeader[25].replace(".","").replace(",","."))
                del cemigHeader
                
                #Assigning nFatura tu "Globalizada" because it represents a lot of bills
                nFatura = "Globalizada"

                #extraindo data de emissão n empresa
                start = ["DATA DE EMISSÃO"]
                end = ["Pague"]
                dataEmi = ((get_keyword(start, end, text)).split(None, 5))[4]

            #Especial para faturas DTCEA-TRM
            else:
                # Identificação
                isNotGlobal = isNotGlobal.split(None, 4)
                ident = isNotGlobal[4]
                del isNotGlobal

                #Getting the Number of the Nota fiscal
                start = ["NOTA FISCAL Nº"]
                end = [" - "]
                nFatura = get_keyword(start,end,text)
                print(nFatura)

                # obtain mes, valor Líquido e data vencimento #3
                start = ["Automático"]
                end = ["_"]
                cemigHeader= get_keyword(start, end, text)
                tempHolder = cemigHeader.split(None, 11)
                dataVenc = tempHolder[7]                
                if(len(tempHolder) > 10):
                    valorLiq = float(tempHolder[9].replace(".","").replace(",","."))
                    tempHolder = (tempHolder[10]).split("/",2)
                    tempHolder[0] = ((tempHolder[0])[0:3]).upper()
                    mes = tempHolder[0]+tempHolder[1]
                else:
                    valorLiq = float(tempHolder[8].replace("R$", "").replace(".","").replace(",","."))
                    tempHolder = (tempHolder[9]).split("/",2)
                    tempHolder[0] = ((tempHolder[0])[0:3]).upper()
                    mes = tempHolder[0]+"/"+tempHolder[1]
                del cemigHeader
                del tempHolder

                #Obtain DATA EMISSÃO
                start = ["emissão:"]
                end = ["Consulte"]
                dataEmi = (get_keyword(start,end,text)).strip()


            #Creating a empty list to fill with all the taxes found in the bill
            #Those search are common for both
            taxesList = []
            start = ["CSLL"]
            end = ["Imposto"]
            taxesList.append(get_keyword(start,end,text))
            start = ["- COFINS"]
            end = ["Imposto"]
            taxesList.append(get_keyword(start,end,text))
            start = ["- PIS/PASEP"]
            end = ["Imposto"]
            taxesList.append(get_keyword(start,end,text))
            start = ["IRPJ"]
            end = ["R$"]
            taxesList.append(((get_keyword(start,end,text)).split(None, 2))[0])
            #Transforming the brazilian notation into american float, and summing to the total value
            darf = 0
            for i in range(0,4):
                taxesList[i] = float(taxesList[i].replace("-","").replace(".","").replace(",","."))
                darf += taxesList[i]

            #defining the total value with taxes

            valorBru = darf+valorLiq

            #obtain Taxa Iluminação da Globaliza e das menores
            start = ["Publica Municipal"]
            end = ["Imposto"]
            tempHolder = get_keyword(start, end, text)
            if((tempHolder is None) or len(tempHolder) == 0):
                iluPlub = ""
            else:
                iluPlub = float(((tempHolder.split(None, 1))[0]).replace(".", "").replace(",","."))
                

            # create a list with the keywords extracted from current document.
            if(nFatura == PDFname):
                for i in range(10):
                    my_list.append("")
                my_list.insert(2, nFatura)
                processFailed = True
            else:
                entregaACI = ""
                my_list = [ident,nFatura,mes,dataEmi,dataVenc,entregaACI,iluPlub,valorLiq, darf, valorBru]
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
            0: "Identificador",
            1: "N° Fatura",
            2: "Mês",
            3: "Data Emissão",
            4: "Data Vencimento",
            5: "Entrega ACI",
            6: "Iluminação Pública",
            7: "Valor Líquido",
            8: "DARF",
            9: "Valor Bruto",
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
