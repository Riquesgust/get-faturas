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
            
            # Caso não seja fatura Globalizada executa esse trecho
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
                
                
                #
                if(((nFatura is None) or (len(nFatura)==0))):
                    nFatura = PDFname
                else:

                    #pegando mês e identificação não empresa
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
                        dataVenc = y[3]
                        del y
                        del tempHolder
                    else:

                        #extraindo identificação e mês
                        y = tempHolder.split(None, 7)
                        mes = y[2]
                        ident = y[6]
                        dataVenc = y[3]
                        del y
                        del tempHolder

                    #extraindo data de emissão n empresa
                    start = ["Data de Emissão: "]
                    end = ["Apresentação"]
                    tempHolder = get_keyword(start, end, text)
                    y = tempHolder.split(None, 4)
                    dataEmi = y[0]
                    del tempHolder
                    del y

                     # obtain Data de emissão n empresa
                    start = ["Data de Emissão: "]
                    end = ["Apresentação"]
                    tempHolder = get_keyword(start, end, text)
                    y = tempHolder.split(None, 4)
                    dataEmi = y[0]
                    del tempHolder
                    del y

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

                # obtain mes, valor Líquido e data vencimento #3
                start = ["Referente a"]
                end = ["BEIRA"]
                cemigHeader= get_keyword(start, end, text)
                cemigHeader= cemigHeader.split(None, 13)
                dataVenc = cemigHeader[12]
                mes = cemigHeader[11]
                valorLiq = float(cemigHeader[13].replace(".","").replace(",","."))
                del cemigHeader

                #Obtain DATA EMISSÃO
                start = ["emissão:"]
                end = ["Consulte"]
                dataEmi = get_keyword(start,end,text)
                dataEmi = dataEmi.strip()

                #Creating a empty list to fill with all the taxes found in the bill
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
                end = ["TOTAL"]
                taxesList.append((get_keyword(start,end,text)))
                start = ["DIC mensal - "]
                end = ["Imposto"]
                #Transforming the brazilian notation into american float, and summing to the total value
                darf = 0
                for i in range(0,4):
                    taxesList[i] = float(taxesList[i].replace("-","").replace(".","").replace(",","."))
                    darf += taxesList[i]

                #defining the total value with taxes
                valorBru = darf+valorLiq

                
                #obtain Taxa Iluminação EMPRESA
                start = ["Total Devoluções/Ajustes"]
                end = ["Consumo"]
                tempHolder = get_keyword(start, end, text)
                if((tempHolder is None) or len(tempHolder) == 0):
                    iluPlub = ""
                else:
                    y = tempHolder.split(None, 4)
                    iluPlub = y[0]
                    if(iluPlub != "10,74"):
                        tempHolder = iluPlub.split("-", 1)
                        iluPlub = float(tempHolder[0].replace(".", "").replace(",", "."))
                    else:
                        iluPlub = float(iluPlub.replace(".", "").replace(",", "."))
        
                    del tempHolder
                    del y
                

            # create a list with the keywords extracted from current document.
            if(nFatura == PDFname):
                for i in range(10):
                    my_list.append("")
                my_list.insert(2, nFatura)
                processFailed = True
            else:
                """
                # obter Valor Líquido #2
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
                    valorLiq = y[0]
                    del tempHolder
                    del y
                else:
                    y = tempHolder.split(None, 4)
                    valorLiq = y[0]
                    del tempHolder
                    del y


                # Calculating correctly the value of "valorBru" if iluPlub different than 10,74
                if((iluPlub != 10.74) and (iluPlub != "")):
                    valorBru = (float(valorBru.replace(".", "").replace(",","."))) - iluPlub
                    darf = valorBru - (float(valorLiq.replace(".", "").replace(",", ".")))
                # Calculating correctly the value of "valorBru" if "iluPlub" is 10,74
                elif((iluPlub == 10.74)):
                    valorBru = (float(valorBru.replace(".", "").replace(",","."))) + iluPlub
                    darf = valorBru - (float(valorLiq.replace(".", "").replace(",", ".")))
                    darf = round(darf, 2)
                else:
                    darf = float(valorBru.replace(".", "").replace(",",".")) - float(valorLiq.replace(".", "").replace(",", "."))
                    darf = round(darf, 2)
                    """
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
