import re                   #importa tudo
from re import search       #importa somete o search
import pandas as pd
import datetime
import os
import win32com.client
from datetime import datetime
import numpy
import time

###################################################################################################################
#EMAIL

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) 
messages = inbox.Items

def saveattachemnts(subject):
    for message in messages:
        if message.Subject == subject:# and message.Unread:
            # body_content = message.body
            attachments = message.Attachments
            attachment = attachments.Item(1)
            for attachment in message.Attachments:
                print(attachment)
                attachment.SaveAsFile(os.path.join("C:\\Users\\VITOVIC\\Documents\\Arquivos-em-Pyhton\\TI VERDE (Control Plan)", str(attachment)))
            if message.Subject == subject and message.Unread:
                    message.Unread = False

if __name__ == "__main__":
    saveattachemnts('Planilhas')#####CAMPO DO ASSUNTO "PLANILHAS"

###################################################################################################################
#TABELAS(DF)
arr = os.listdir(r'C:\Users\VITOVIC\Documents\Arquivos-em-Pyhton\TI VERDE (Control Plan)')
#print(arr) printa a lista dentro dos documentos 

contCert=1
contPag=1
for x in arr:
    if re.search("CERTIFICADO",str(x)):
        os.rename('C:\\Users\\VITOVIC\Documents\\Arquivos-em-Pyhton\\TI VERDE (Control Plan)'+
        "\\"+x,'C:\\Users\\VITOVIC\\Documents\\Arquivos-em-Pyhton\\TI VERDE (Control Plan)'+
        "\\"+'CERTIFICADO'+str(contCert)+'.xlsx')
        print('Encontrei um Certificado')
        contCert+=1
    if re.search("PAGAMENTO",str(x)):
        os.rename('C:\\Users\\VITOVIC\\Documents\\Arquivos-em-Pyhton\\TI VERDE (Control Plan)'+
        "\\"+x,'C:\\Users\\VITOVIC\\Documents\\Arquivos-em-Pyhton\\TI VERDE (Control Plan)'+
        "\\"+'PAGAMENTO'+str(contPag)+'.xlsx')
        print("Encontrou um Pagamento")
        contPag+=1
contCert-=1
        
for x in (range(contCert)):
    tabela1 = ("C:\\Users\\VITOVIC\\Documents\\Arquivos-em-Pyhton\\TI VERDE (Control Plan)\\CERTIFICADO"+str(x+1)+".xlsx")
    df = pd.read_excel(tabela1)

    tabela2 = ("C:\\Users\\VITOVIC\\Documents\\Arquivos-em-Pyhton\\TI VERDE (Control Plan)\\PAGAMENTO"+str(x+1)+".xlsx")
    df2 =pd.read_excel(tabela2)

    dados_consolidados = r"C:\Users\VITOVIC\Documents\Arquivos-em-Pyhton\TI VERDE (Control Plan)\CONTROLE\Dados consolidados TI Verde.xlsx"
    df3 = pd.read_excel(dados_consolidados,header=None)

    ###################################################################################################################
    #TIPOS EXISTENTENTES DE LIXO (QUE TEMOS COMO BASE NA TABELA)

    tipoLixo=['LIXO TECNOLÓGICO (TERMINAIS)',
    'LIXO TECNOLÓGICO ', 'PILHAS E BATERIAS ', 'RESÍDUO DE AGÊNCIAS NÃO REEE',
    'PAPEL', 'LIXO TECNOLOGICO', 'LIXO TECNOLOGICO(TERMI/ACESSO.)',
    'RESÍDUO METÁLICO', 'LIXO TECNOLOGICO(T.AS/ATMS/URNAS)',
    'LIXO TECNOLOGICO (TCR VERTERA)', 'BATERIAS ACIDAS',
    'LIXO TECNOLOGICO (T.AS/ATMS/URNAS)', 'LIXO TECNOLOGICO (TERMI/ACESSO.)',
    'BATERIAS (NOBREAK)', 'LIXO TECNOLOGICO (ATMS)', 'LIXO TECNOLOGICO (COFRE)',
    'LIXO TECNOLOGICO (MOB. DIVERSOS)']
    #Refinar com ela

    ###################################################################################################################

    '''
    fORMA COMO ACHEI A QUANTIDADE UNICA DE CADA ITEM DENTRO DA COLUNA
    TIPOS DE RESIDUO QUE TEMOS COMO BASE NA PLANILHA
    print(df3)
    listaTipo=df3[8].unique()
    print(str(listaTipo))
    '''
    ###################################################################################################################
    #Valor do item
    listaValor=df2['Unnamed: 3'].to_list()
    #print(listaValor) # <<<<<<CASO queira confirmar todos os valores existentes
    for i in (i for i, x in enumerate(listaValor)if str(x) == 'nan'):
        item=i-2
        break
    #REFINAR COM ELA A FORMA DA TABELA DE VALOR
    #print(listaValor[item])<<<<<CASO queira ver todos os itens dentro de valor
    valor=listaValor[item]


    #Flags <<<usados para busca de alguns itens
    iDataDoc = 0
    iDataChegada = 0
    iQuant = 0
    iTipo = 0

    ###################################################################################################################
    #TRANSFORMANDO O DF INTEIRO EM LISTA 
    lista = df.values.tolist()

    #COMEÇA A BUSCA DOS ITEMS DENTRO DO DF(LISTA).

    for item in lista:
    ####CERTIFICADO
        if re.search("\s[0-9]{2}\.[0-9]{3}\/[0-9]{2}",str(item))is not None:
            local = re.search('\s[0-9]{2}\.[0-9]{3}\/[0-9]{2}', str(item)).span()
            certificado=str(item)[local[0]:local[1]]   

        if re.search("Data Doc",str(item))is not None:
            iDataDoc = 1

        if re.search("Quantidade",str(item))is not None:
            iQuant = 1
            item = [str(i) for i in item]

        if re.search("Tipo Resíduo",str(item))is not None:
            iTipo = 1
            item = [str(i) for i in item]
        if re.search("Data Chegada",str(item))is not None:
            iDataChegada = 1
            item = [str(i) for i in item]
    ####DATA DOCUMENTO(DOC)
        if re.search("[0-9]{2}\/[0-9]{2}\/[0-9]{4}",str(item))is not None:
            if iDataDoc == 1:
                local = re.search('[0-9]{2}\/[0-9]{2}\/[0-9]{4}', str(item)).span()
                dataDoc=str(item)[local[0]:local[1]]
                iDataDoc = 0
    ####DATA CHEGADA 
        if re.search("[0-9]{4}\-[0-9]{2}\-[0-9]{2}",str(item))is not None:
            if iDataChegada == 1:
                local = re.search('[0-9]{4}\-[0-9]{2}\-[0-9]{2}',str(item)).span()
                datachegada=str(item)[local[0]:local[1]]
                iDataChegada = 0 

    ####NUMERO DOCUMENTO(DOC)
        if re.search("[0-9]{3}\/[0-9]{4}\-[A-Z]{3}",str(item))is not None:
                local = re.search('[0-9]{3}\/[0-9]{4}\-[A-Z]{3}', str(item)).span()
                numeroDoc=str(item)[local[0]:local[1]]           
    ####PESO TOTAL   
        if re.search("[0-9]*\s[A-Z]{2}",str(item))is not None:
            if iQuant == 1:
                local = re.search('[0-9]*\s[A-Z]{2}', str(item)).span()
                pesoTotal=str(item)[local[0]:local[1]]
                iQuant= 0 

    ####Tipo Residuo
    ####Fazendo com regex #str('|'.join(map(str, lista2)))
        for i in item:
            if i in tipoLixo:
                tipoResiduo = i
    ####PERGUNTAR SE O PADRAO DO TIPO DE LIXO É SEMPRE EM UPPER CASE
    ###################################################################################################################
    #CONFERINDO OS ITENS
    print("Certificado:"+certificado)
    print("Data documento:"+dataDoc)
    print("Data chegada:"+datachegada)
    print("Numero do documento:"+numeroDoc)
    print("Peso Total:"+pesoTotal)
    print("Tipo de Residuo:"+tipoResiduo)

    dataconvertida = datetime.strptime(dataDoc, '%d/%m/%Y')

    dia=dataconvertida.strftime('%d')###dia
    mes=dataconvertida.strftime('%m')###mes
    ano=dataconvertida.strftime('%y')###ano

    ###################################################################################################################
    ###################################################################################################################
    print("Salvo na Planilha!")
    df3 = df3.append(pd.DataFrame([['','',numeroDoc,dataDoc,pesoTotal,valor,dataDoc,
        '',tipoResiduo,certificado,dataDoc,pesoTotal,valor,'',dia,mes,ano,dataDoc]]),ignore_index=True)


    writer=pd.ExcelWriter(r"C:\Users\VITOVIC\Documents\Arquivos-em-Pyhton\TI VERDE (Control Plan)\CONTROLE\Dados consolidados TI Verde.xlsx",engine='xlsxwriter')
    workbook = writer.book
    df3.to_excel(writer,sheet_name='Dados Controle Certificado',index=False,header=False)
    worksheet = writer.sheets['Dados Controle Certificado']

    cell_format = workbook.add_format()
    cell_format2 =workbook.add_format()
    cell_format.set_bold(True)
    cell_format.set_border(1) #<<<<<<<<<<<<<ativa o negrito da das letras (False) <<<para nao desativa o negrito
    cell_format2.set_pattern(1)#Trabalha com a cor de fundo vai de 1 a 18 (1 é otimo 10 é pessimo 18 razuavel)
    cell_format2.set_bg_color('orange')
    cell_format2.set_font_name('Arial Black')

    worksheet.write('B2','ORIGEM',cell_format2) #Começa na B2
    worksheet.write('C2','Nº da Proposta',cell_format2)
    worksheet.write('D2','Data do Documento',cell_format2)
    worksheet.write('E2','Peso Propostas [kg]',cell_format2)
    worksheet.write('F2','Valor da Proposta[R$]',cell_format2)
    worksheet.write('G2','Data Depósito',cell_format2)
    worksheet.write('H2','Enviado para Baixa do Imobilizado em:',cell_format2)
    worksheet.write('I3','Tipo Residuo',cell_format2)
    worksheet.write('J3','Nº Numero',cell_format2)
    worksheet.write('K3','Data Certificado',cell_format2)
    worksheet.write('L3','Peso',cell_format2)
    worksheet.write('M3','Valor[R$]',cell_format2)
    worksheet.write('N3','',cell_format2)
    worksheet.write('O3','Dia',cell_format2)
    worksheet.write('P3','Mes',cell_format2)
    worksheet.write('Q3','Ano',cell_format2)
    worksheet.write('R3','Data Envio',cell_format2,)

    cell_format3 = workbook.add_format()
    cell_format3 = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'blue'})
    worksheet.merge_range('I2:R2', 'Certificado', cell_format3)
    cell_format3.set_font_name('Arial Black')

    cell_format.set_font_color('black')
    cell_format2.set_font_color('black')
    worksheet.set_column('B:R',20, cell_format)#<<<<<<<<<coluna

    writer.save()


arr = os.listdir(r'C:\Users\VITOVIC\Documents\Arquivos-em-Pyhton\TI VERDE (Control Plan)')
#print(arr)
print("TODOS OS ARQUIVOS EXTRAIDOS SALVOS NA PLANILHA")
tempo=0
for x in arr:
    dataHoje=datetime.now()
    dataHojeTexto=(dataHoje.strftime('%Y-%m-%d %H-%M-%S'))
    if re.search("CERTIFICADO",str(x)):
        os.rename(r'C:\Users\VITOVIC\Documents\Arquivos-em-Pyhton\TI VERDE (Control Plan)'+
        "\\"+x,r"C:\Users\VITOVIC\Documents\Arquivos-em-Pyhton\TI VERDE (Control Plan)\PLAN(Ja verificadas)"+
        "\\"+'CERTIFICADO(verificado'+dataHojeTexto+').xlsx')
        time.sleep(2)
    if re.search("PAGAMENTO",str(x)):
        os.rename(r'C:\Users\VITOVIC\Documents\Arquivos-em-Pyhton\TI VERDE (Control Plan)'+
        "\\"+x,r"C:\Users\VITOVIC\Documents\Arquivos-em-Pyhton\TI VERDE (Control Plan)\PLAN(Ja verificadas)"+
        "\\"+'PAGAMENTO(Verificado'+dataHojeTexto+').xlsx')
        time.sleep(2)

        #C:\Users\VITOVIC\Documents\Arquivos-em-Pyhton\TI VERDE (Control Plan)\PLAN(Ja verificadas)
