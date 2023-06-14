import csv
import os.path
import re
import tkinter as tk
import tkinter.ttk as ttk
from statistics import mode
from tkinter import *
from tkinter import filedialog

import ocrmypdf
from pypdf import PdfReader

appName = "AgroPontos RegEx v1.4"

#Abre janela do windows para seleção do arquivo PDF/TXT
def pick_file(self):
    filePath = filedialog.askopenfilename()
    if filePath:
        self.fileLabel.config(text=filePath)
    else:
        self.fileLabel.config(text="Arquivo não existe ou não foi selecionado um arquivo.")
    return filePath

#Mostra as mensagens na interface
def show_text(sText, self):
    self.outputLabel.config(text=sText)
    self.text.delete("1.0",END)
    self.text.insert(END,self.outputLabel.cget("text"))
    self.window.update()

#Calcula a moda das coordenadas lidas, para identificar falhas de leitura
def mode_len(lst):
    lengths = [len(elem) for elem in lst]
    return mode(lengths)

#Remove pontos extras, caso haja múltiplos devido a notação brasileira ou falhas de leitura
def remove_extra_dot(s):
    lastDotIndex = s.rfind('.')
    if lastDotIndex == -1:
        return s
    else:
        return s[:lastDotIndex].replace('.', '') + s[lastDotIndex:]

#Utiliza a biblioteca OCRMYPDF para ler PDFs escaneados e consertá-los (desvira páginas, corrige angulação...)
#Ou se o PDF já contém texto, utiliza a biblioteca pypdf para extraí-lo
def process_pdf(self):
    filePath = self.fileLabel.cget("text")
    txtFilePath = os.path.splitext(filePath)[0]  + '.txt'
    
    if not os.path.exists(filePath):
        show_text("Arquivo não encontrado: "+filePath, self)
        return None

    if self.varText.get() == 1:
        show_text("Extraindo texto do PDF...", self)
        textPDF = ''
        reader = PdfReader(filePath)
        
        for page in reader.pages:
            textPDF = textPDF + str(page.extract_text(0))
        with open(txtFilePath, 'w', encoding='UTF-8') as f:
            f.write(textPDF)
        show_text("Texto extraído com sucesso!", self)
        return
    else:
        show_text("Processando PDF... Aguarde, isso pode levar alguns minutos.", self)
        if __name__ == '__main__':  # Da biblioteca, para funcionar direito em Windows e Mac
            ocrmypdf.ocr(filePath,os.path.splitext(filePath)[0]+'_NOVO.pdf', deskew=True, clean=True, rotate_pages=True, force_ocr=True, language='por', sidecar=txtFilePath)
        show_text("PDF processado com sucesso!", self)

#Exporta as coordenadas encontradas para um arquivo CSV
def export_csv(self):
    #matches = match_string()
    matches = self.text.get("2.0","end-1c")
    splitChar = self.splitEntry.get()

    if matches is None:
        return
    
    matchesAux = matches.split("\n")
    matches = []
    
    numChars = int(self.num_chars_spinbox.get())
    if numChars > 0:
        for matchAux in matchesAux:
            matches.append(matchAux[:-numChars])
    else:
        matches = matchesAux

    csvFilePath = os.path.splitext(self.fileLabel.cget("text"))[0] + '.csv'

    if self.coordType.get() == "UTM":
        with open(csvFilePath, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow('NE')
            NList = []
            EList = []
            errorList = ''
            for match in matches:
                if self.varDecimal.get() == 0:
                    matchAux = match.replace('.','').replace(',','.')
                else:
                    matchAux = match.replace(',','')
                matchAux = re.sub(r'[^0-9.'+splitChar+']', '', matchAux)
                [N, E] = matchAux.split(splitChar)
                N = remove_extra_dot(N)
                E = remove_extra_dot(E)
                NList.append(N)
                EList.append(E)
                if (len(N) != mode_len(NList) or len(E) != mode_len(EList)):
                    errorList = errorList + match + '\n'
                writer.writerow([N, E])
            if errorList:
                show_text('Coversão PDF -> TXT substituiu algum número da coordenada por letra, ou há caracteres faltando.\nClique novamente em "Encontrar" e edite-as na interface, ou diretamente no CSV.\n'+errorList, self)
                print("Coordenadas com Erro: \n" + errorList)

    if self.coordType.get() == "Lat-Long":
        with open(csvFilePath, 'w', newline='') as f:
            writer = csv.writer(f, quotechar = "@")
            writer.writerow(["NOME","LONG","LAT"])
            nameList = []
            longList = []
            latList = []
            errorList = ''
            for match in matches:
                if self.varDecimal.get() == 0:
                    matchAux = match.replace('.','').replace(',','.')
                else:
                    matchAux = match.replace(',','')
                [cName, cLong, cLat] = matchAux.split(splitChar)
                cLong = remove_extra_dot(cLong)
                cLat = remove_extra_dot(cLat)
                nameList.append(cName)
                longList.append(cLong)
                latList.append(cLat)
                if (len(cLong) != mode_len(longList) or len(cLat) != mode_len(latList)):
                    errorList = errorList + match + '\n'
                writer.writerow([cName, cLong, cLat])
            if errorList:
                show_text('Coversão PDF -> TXT substituiu algum número da coordenada por letra, ou há caracteres faltando.\nClique novamente em "Encontrar" e edite-as na interface, ou diretamente no CSV.\n'+errorList, self)
                print("Coordenadas com Erro: \n" + errorList)

#Encontra as coordenadas no arquivo TXT, utilizando-se de regras de Regex (https://docs.python.org/3/library/re.html)
def match_string(self):
    txtFilePath = os.path.splitext(self.fileLabel.cget("text"))[0] + '.txt'
    self.window.update()
    # print (window.winfo_width())
    # print (window.winfo_height())
    # print (window.winfo_geometry())

    if not os.path.exists(txtFilePath):
        show_text("Arquivo não encontrado: "+txtFilePath, self)
        return None
    
    with open(txtFilePath, 'r', encoding='UTF-8') as txtFile:
        txtData = txtFile.read()      
        
    startPattern = self.startEntry.get()
    endPattern = self.endEntry.get()

    string = ''
    
    if startPattern[0].isalpha() or startPattern[0]=='\\':
        if endPattern[0].isalpha() or endPattern[0]=='\\':
            pattern = startPattern + '.{10,100}?' + endPattern
        else:
            pattern = startPattern + '.{10,100}?\\' + endPattern
    else:
        if endPattern[0].isalpha() or endPattern[0]=='\\':
            pattern = '\\' + startPattern + '.{10,100}?' + endPattern
        else:
            pattern = '\\' + startPattern + '.{10,100}?\\' + endPattern

    matchesAux = []
    matches = re.findall(pattern, txtData, flags=re.DOTALL)
    for match in matches:
        if match == matches[-1]:
            match = re.sub(r"[\r\n]+", " ", match)
            string = string + match
            matchesAux.append(match)
        else:
            match = re.sub(r"[\r\n]+", " ", match)
            string = string + match + '\n'
            matchesAux.append(match)
    if matches:
        string = re.sub(' {2,}', ' ',string)
        show_text("Coordenadas encontradas: "+str(len(matches))+"\n"+string, self)
        return matchesAux
    else:
        show_text("O padrão especificado não foi encontrado.", self)
        return None

class GUI():
    def __init__(self):
        super().__init__()

        #Cria a janela principal
        self.window = tk.Tk()
        self.window.title(appName)
        self.window.geometry("665x697")
        self.window.configure(background="white")
        self.window.resizable(0,0)
        self.window.iconbitmap("favicon.ico")

        #Define estilo para os elementos gráficos
        self.style = ttk.Style()
        #self.style.theme_use('clam')
        self.style.configure('TButton', font=('Helvetica', 10), background='white', foreground='black')
        self.style.configure('TLabel', font=('Helvetica', 10), background='white', foreground='black')
        self.style.configure('TEntry', font=('Helvetica', 10), background='white', foreground='black')
        self.style.configure('TCheckbutton', font=('Helvetica', 10), background='white', foreground='black')
        self.style.configure('TOptionMenu', font=('Helvetica', 10), background='white', foreground='black')

        #Cria 3 frames na janela
        self.frameTop = tk.Frame(self.window, bg="white")
        self.frameMiddle = tk.Frame(self.window, bg="white")
        self.frameBottom = tk.Frame(self.window, bg="white")

        #Cria e adiciona elementos ao frame superior
        self.fileLabel = ttk.Label(self.frameTop, text="Selecione o arquivo PDF/TXT:")
        self.fileLabel.pack(side="top")
        self.fileButton = tk.Button(self.frameTop, text="Selecionar", command=lambda: pick_file(self))
        self.fileButton.pack(side="top")

        self.startLabel = ttk.Label(self.frameTop, text="Digite o início do padrão:")
        self.startLabel.pack(anchor='c')
        self.startEntry = ttk.Entry(self.frameTop)
        self.startEntry.pack(anchor='c')

        self.endLabel = ttk.Label(self.frameTop, text="Digite o final do padrão:")
        self.endLabel.pack(anchor='c')
        self.endEntry = ttk.Entry(self.frameTop)
        self.endEntry.pack(anchor='c')

        self.splitLabel = ttk.Label(self.frameTop, text="Digite o caractere de separação:")
        self.splitLabel.pack()
        self.splitEntry = ttk.Entry(self.frameTop)
        self.splitEntry.pack()

        self.num_chars_label = ttk.Label(self.frameTop, text="Qtd de caracteres para remover do fim:")
        self.num_chars_label.pack()
        self.numCharsDefault = StringVar(self.frameTop)
        self.numCharsDefault.set('0')
        self.num_chars_spinbox = ttk.Spinbox(self.frameTop, from_=0, to=50, width=5, textvariable=self.numCharsDefault)
        self.num_chars_spinbox.pack()

        self.varText = tk.IntVar()
        self.checkText = ttk.Checkbutton(self.frameTop, text='PDF é Texto',variable=self.varText, onvalue=1, offvalue=0)
        self.checkText.pack(anchor="sw",side="left")

        self.varDecimal = tk.IntVar()
        self.checkDecimal = ttk.Checkbutton(self.frameTop, text='Decimal Americano',variable=self.varDecimal, onvalue=1, offvalue=0)
        self.checkDecimal.pack(anchor="sw",side="left")

        self.coordType = StringVar(self.frameTop)
        self.coordValues = ["","UTM", "Lat-Long"]
        self.coordType.set(self.coordValues[1])

        self.coordOption = ttk.OptionMenu(self.frameTop, self.coordType, *self.coordValues)
        self.coordOption.pack(anchor="se",side="right")

        #Cria e adiciona elementos ao frame intermediário (saída de dados)
        self.outputFrame = ttk.Frame(self.frameMiddle)
        self.outputLabel = ttk.Label(self.outputFrame, wraplength=500, justify="left")
        self.outputLabel.pack(side="left")
        self.text = tk.Text(self.outputLabel)
        self.text.grid(row=0, column=1)

        self.yscrollbar = ttk.Scrollbar(self.outputFrame, orient="vertical", command=self.text.yview)
        self.yscrollbar.pack(side="right", fill="y")
        self.outputFrame.pack(fill="both", expand=True)
        self.text.config(yscrollcommand=self.yscrollbar.set)

        #Cria e adiciona elementos ao frame inferior
        self.hoverLabel = ttk.Label(self.frameBottom, text="Informação")
        self.hoverLabel.pack(anchor="sw",side="bottom")

        self.matchButton = ttk.Button(self.frameBottom, text="Processar PDF", command=lambda: process_pdf(self))
        self.matchButton.pack(side="top")

        ttk.Label(self.frameBottom, text="").pack(side="right", fill="both", expand=True)

        self.csvButton = ttk.Button(self.frameBottom, text="Gerar CSV", command=lambda: export_csv(self))
        self.csvButton.pack(side="right")

        ttk.Label(self.frameBottom, text="").pack(side="left", fill="both", expand=True)

        self.matchButton = ttk.Button(self.frameBottom, text="Encontrar", command=lambda: match_string(self))
        self.matchButton.pack(side="left")

        #Cria e adiciona o tooltip com informações
        self.tooltip = ttk.Label(self.window, text="Use [] para fazer um OU(lógico) dos elementos. \n Exemplo: Utilize )[,.] para encontrar strings que contenham ), ou ). ")
        self.tooltip.configure(background="#f0f0f0", foreground="black", font=("Helvetica", 10))
        self.tooltip.bind("<Enter>", lambda event: self.tooltip.place(relx=0.5, rely=0.5, anchor="c"))
        self.tooltip.bind("<Leave>", lambda event: self.tooltip.place_forget())

        self.hoverLabel.bind("<Enter>", lambda event: self.tooltip.place(relx=0.5, rely=0.5, anchor="c"))
        self.hoverLabel.bind("<Leave>", lambda event: self.tooltip.place_forget())

        #Define e posiciona os 3 frames na janela
        self.frameTop.pack(side="top",fill="both", expand=True)
        self.frameMiddle.pack(fill="both", expand=True)
        self.frameBottom.pack(side="bottom",fill="both", expand=True)

        self.window.mainloop()

def main():
    GUI()

if __name__ == "__main__":
    main()