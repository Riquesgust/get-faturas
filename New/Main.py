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
    #saves the path of the file
    currentPath = str(filedialog.askdirectory())    
    #add the *.pdf fucntion to open all the files in the directory
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
            
            # distinguishing between Residencial e Empresa
            start = ["www.cpflempresas.com.br"]
            end = ["Série"]
            isResidencial = get_keyword(start, end, text)

            # If it isn't company
            if((isResidencial is None) or (len(isResidencial)==0)):
                #If it is Residencial assigns the Iluminação publica variable to 0
                iluPlub = ""
                
                # obtain Número da Fatura Residencial
                start = ["Nº"]
                end = ["Série"]
                nFatura = get_keyword(start, end, text)
                if(((nFatura is None) or (len(nFatura)==0))):
                    nFatura = PDFname
                else:

                    #obtain mês and identificação residencial
                    start = ["0800 010 2570"]
                    end = ["Descrição"]
                    tempHolder = get_keyword(start, end, text)

                    #if does not find the first telephone number search for the other
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

                        #obtain identificação and mês
                        y = tempHolder.split(None, 7)
                        mes = y[2]
                        ident = y[6]
                        dataVenc = y[3]
                        del y
                        del tempHolder

                    #obtain data de emissão Residencial
                    start = ["Data de Emissão: "]
                    end = ["Apresentação"]
                    tempHolder = get_keyword(start, end, text)
                    y = tempHolder.split(None, 4)
                    dataEmi = y[0]
                    del tempHolder
                    del y

                     # obtain Data de emissão Residencial
                    start = ["Data de Emissão: "]
                    end = ["Apresentação"]
                    tempHolder = get_keyword(start, end, text)
                    y = tempHolder.split(None, 4)
                    dataEmi = y[0]
                    del tempHolder
                    del y

            #Made for company bills
            else:
                # obtain Número da Fatura #1
                start = ["Nº."]
                end = ["série"]
                nFatura = get_keyword(start, end, text)

                # obtain mes, n° serie and data vencimento EMPRESA #3
                start = ["0800 770 4140"]
                end = ["Descrição"]
                tempHolder = get_keyword(start, end, text)
                y = tempHolder.split(None, 5)
                mes = y[2]
                ident = y[1]
                dataVenc = y[3]
                tempList = list(mes.split("/"))
                tempList1 = tempList[1][2:]
                y = tempList[0]+'/'+tempList1
                del tempHolder
                del tempList1
                del tempList

                # obtain Data de emissão empresas
                start = ["Data de Emissão "]
                end = ["Apresentação"]
                tempHolder = get_keyword(start, end, text)
                y = tempHolder.split(None, 4)
                dataEmi = y[0]
                del tempHolder
                del y

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

            # create a list with the keywords extracted from current document
            if(nFatura == PDFname):
                for i in range(10):
                    my_list.append("")
                my_list.insert(2, nFatura)
                processFailed = True
            else:
                # obtain Valor Líquido #2
                start = ["Total Consolidado"]
                end = ["Consumo"]
                tempHolder = get_keyword(start, end, text)
                
                #If the value was not found it will search in the next page
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

               

                # obtain Data de emissão #4
                start = ["Total Distribuidora"]
                end = ["Consumo"]
                tempHolder = get_keyword(start, end, text)
                y = tempHolder.split(None, 3)
                valorBru = y[0]
                del tempHolder
                del y
                # correctly calculating the value of "valorBru" if iluPlub different than 10,74
                if((iluPlub != 10.74) and (iluPlub != "")):
                    valorBru = (float(valorBru.replace(".", "").replace(",","."))) - iluPlub
                    darf = valorBru - (float(valorLiq.replace(".", "").replace(",", ".")))
                # correctly calculating  the value of "valorBru" if "iluPlub" is 10,74
                elif((iluPlub == 10.74)):
                    valorBru = (float(valorBru.replace(".", "").replace(",","."))) + iluPlub
                    darf = valorBru - (float(valorLiq.replace(".", "").replace(",", ".")))
                    darf = round(darf, 2)
                else:
                    darf = float(valorBru.replace(".", "").replace(",",".")) - float(valorLiq.replace(".", "").replace(",", "."))
                    darf = round(darf, 2)
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
