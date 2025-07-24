import pymupdf
import streamlit as st
import zipfile
import os
import numpy as np
import pandas as pd
import random
import subprocess
import datetime
from PyPDF2 import PdfReader, PdfWriter
    
@st.cache_data   
def nameFile():
    symbols = ['-', ':', '.']
    nowTime = str(datetime.datetime.now())
    try:
        for symbol in symbols: 
            nowTime = nowTime.replace(symbol, '_')
    except:
        pass
    return nowTime
    
@st.cache_data  
def extractText(filePdf):
    text = ''
    docPdf = pymupdf.open(filePdf)
    for page in docPdf:
       text += page.get_text()
    docPdf.close()
    return text
        
@st.cache_data
def extractUrls(filePdf):
    docPdf = pymupdf.open(filePdf)
    allLinks = []
    for p, page in enumerate(docPdf):
        links = page.get_links()
        for link in links:
            nameUrl = link["uri"]
            fromUrl = link["from"]
            newText = f'{nameUrl}; coordenadas: {fromUrl}\n'
            allLinks.append(newText)
    text = ''.join(allLinks) 
    docPdf.close()
    return text
    
def mensResult(value, nFiles, modelButt, fileTmp, fileFinal):
    colMens, colDown = st.columns([8, 2]) 
    if value == 1:
        if modelButt == 'zip': 
                with open(fileTmp, "rb") as file:
                    colDown.download_button(label="Download",
                                            data=file,
                                            file_name=fileFinal,
                                            mime='application/zip', 
                                            icon=":material/download:", 
                                            use_container_width=True)
        colMens.success(f'Gerado o zipado :blue[**{fileFinal}**] com ***{nFiles}*** arquivo(s). Clique no botão ao lado 👉.', icon='✔️') 
    elif value == 0:
        colDown.download_button(label='Download', 
                           data=fileTmp,
                           file_name=fileFinal,
                           mime='application/octet-stream', 
                           icon=":material/download:", 
                           use_container_width=True)
        colMens.success(f'Gerado o arquivo :blue[**{fileFinal}**]. Clique no botão ao lado 👉.', icon='✔️') 
    elif value == 2:
        colDown.download_button(label='Download',
                           data=fileTmp,
                           file_name=fileFinal,
                           mime="text/csv", 
                           icon=":material/download:", 
                           use_container_width=True)
        colMens.success(f'Gerado o arquivo :blue[**{fileFinal}**]. Clique no botão ao lado 👉.', icon='✔️')
    
def extractImgs(filePdf):
    docPdf = pymupdf.open(filePdf)
    allImgName = []
    for p, page in enumerate(docPdf):
        imageList = page.get_images(full=True)
        if imageList:
            for i, img in enumerate(imageList):
                xref = img[0]
                baseImg = docPdf.extract_image(xref)
                imgBytes = baseImg["image"]
                imgExt = baseImg["ext"]
                imgName = f"image_{p+1}_{i+1}.{imgExt}"
                with open(imgName, "wb") as fileImg:
                    fileImg.write(imgBytes)
                allImgName.append(imgName)
    return allImgName

def downloadExt(files, namePdf, numPgOne, numPgTwo, obj):
    fileTmp = f'{nameFile()}_tempFile.zip'
    fileZip = f'namePdf_{numPgOne}_{numPgTwo}_{nameFile()}.zip'
    for file in files:
        with open(file, "rb") as extFile:
           PDFbyte = extFile.read()
        with zipfile.ZipFile(fileTmp, 'a') as extFile:
           extFile.writestr(file, PDFbyte)
    nFiles = len(files) 
    if nFiles > 0:
        mensResult(1, len(files), 'zip', fileTmp, fileZip)
    else:
        strEmpty = f'😢 Extração fracassada!\n🔴 arquivo {namePdf} \nsem {obj} extraível no intervalo de páginas {numPgOne}-{numPgTwo}!'
        config(strEmpty)

def rotatePdf(filePdf, index):
    inputPdf = filePdf
    name, ext = os.path.splitext(inputPdf)
    angle = int(valAngles[index].replace('°', ''))
    output = f'{name}_rotate_{angle}{ext}'
    docPdf = pymupdf.open(filePdf)
    for page in docPdf:
        page.set_rotation(angle)
    docPdf.save(output)
    docPdf.close()
    return output
    
def saveAllPdf(outputBase, contPartes, writer):
    outputPdf = f"{outputBase}{contPartes + 1}.pdf"
    with open(outputPdf, "wb") as outputFile:
        writer.write(outputFile)
    docPdf = pymupdf.open(outputFile)
    countPg.append(len(docPdf))
    docPdf.close()
    return outputPdf

def divideBySize(inputPdf, sizeMax, outputBase):
    filesCutSave = []
    try:
        reader = PdfReader(inputPdf)
        nPgs = len(reader.pages)
        sizeActual = 0
        contPartes = 0
        writer = PdfWriter()
        for i in range(nPgs):
            nameTeste = f'teste_{i+1}.pdf'
            namesTeste.append(nameTeste)
            page = reader.pages[i]
            writer.add_page(page) 
            with open(nameTeste, 'wb') as g:
                writer.write(g)
            sizeActual = os.path.getsize(nameTeste)/(1024**2)
            if sizeActual >= sizeMax:
                outputPdf = saveAllPdf(outputBase, contPartes, writer)
                filesCutSave.append(outputPdf)
                writer = PdfWriter()
                sizeActual = 0
                contPartes += 1
        if len(writer.pages) > 0:
            outputPdf = saveAllPdf(outputBase, contPartes, writer)
            filesCutSave.append(outputPdf)
    except Exception as e:
        st.error(f"Ocorreu um erro: {e} - página {i+1}", icon='🛑')
    return filesCutSave    

def createPdfSel(docPdf, numPgOne, numPgTwo, namePdf, index):
    numPgOne -= 1    
    inputPdf = docPdf
    name, ext = os.path.splitext(namePdf)
    outputPdf = f'{name}_{numPgOne}_{numPgTwo}.pdf'
    listSel = [pg for pg in range(numPgOne, numPgTwo)]
    docPdf.select(listSel)
    docPdf.save(outputPdf)
    if index != 4:
        outputPdf = rotatePdf(outputPdf, index) 
    docPdf.close()
    return outputPdf   

def addWatermark(inputPdf, valMark):
    doc = pymupdf.open(inputPdf)
    for page_num in range(doc.page_count):
        page = doc[page_num]
        page_rect = page.rect
        x = page_rect.width/6.5
        y = page_rect.height * 0.95
        page.insert_text(
            (x, y),  
            valMark,  
            fontsize=25,
            color=(0.7, 0.7, 0.4),  
            rotate=0,  
            fill_opacity=0.3,
            stroke_opacity=0.3
        )
    doc.save(inputPdf, incremental=True, encryption=0)
    doc.close()
    return inputPdf

def selPdfMark(docPdf, numPgOne, numPgTwo, namePdf, index, valMark):
    outputPdf = createPdfSel(docPdf, numPgOne, numPgTwo, namePdf, index)
    pdfMark = addWatermark(outputPdf, valMark)
    downPdfUnique(pdfMark, numPgOne, numPgTwo, namePdf)           
    
def selPgsSize(docPdf, numPgOne, numPgTwo, namePdf, index, sizeMax):
    outputPdf = createPdfSel(docPdf, numPgOne, numPgTwo, namePdf, index)
    inputPdf = outputPdf
    sizeMaxStr = str(sizeMax).replace('.', '_')
    sizeSplit = sizeMaxStr.split('_')
    try:
        numOne = sizeSplit[0]
        numTwo = sizeSplit[1][:2]
        if numTwo.strip() == '00':
             numTwo = ''
        sizeMaxStr = numOne + '_' + numTwo 
    except:
        pass
    outputBase = f'{os.path.splitext(inputPdf)[0]}_divisão_{sizeMaxStr}_Mb__parte_'
    filesCutSave = divideBySize(inputPdf, sizeMax, outputBase)
    downloadExt(filesCutSave, namePdf, numPgOne, numPgTwo, 'pedaços')
        
def selImgUrlsPgs(docPdf, numPgOne, numPgTwo, namePdf, mode, index):
    outputPdf = createPdfSel(docPdf, numPgOne, numPgTwo, namePdf, index)
    filesImg = extractImgs(outputPdf)
    downloadExt(filesImg, namePdf, numPgOne, numPgTwo, 'imagens')
    
def selTxtUrlPgs(docPdf, numPgOne, numPgTwo, namePdf, mode, index):
    outputPdf = createPdfSel(docPdf, numPgOne, numPgTwo, namePdf, index)
    if mode == 0:
        text = extractText(outputPdf)
        strLabel = "Download_text"
        outputTxt = f'{namePdf}_{numPgOne}_{numPgTwo}_text.txt'
        strEmpty = f'😢 Extração fracassada!\n🔴 arquivo {namePdf} \nsem texto extraível no intervalo de páginas {numPgOne}-{numPgTwo}!'
    else:
        text = extractUrls(outputPdf)
        strLabel = "Download_urls"
        outputTxt = f'{namePdf}_{numPgOne}_{numPgTwo}_urls.txt'
        strEmpty = f'😢 Extração fracassada!\n🔴 arquivo {namePdf} \nsem URL extraível no intervalo de páginas {numPgOne}-{numPgTwo}!'
    if len(text.strip()) > 0:
        mensResult(2, 1, 'txt', text, outputTxt)
        #colResult, colDow = st.columns([10, 2])
        #st.download_button(label=strLabel,
        #                   data=text,
        #                   file_name=outputTxt,
        #                   mime="text/csv", 
        #                   icon=":material/download:")
    else:
        config(strEmpty)    
                                          
def selDelPgs(docPdf, numPgOne, numPgTwo, namePdf, mode, index):
    numPgOne -= 1
    inputPdf = docPdf
    name, ext = os.path.splitext(namePdf)
    if mode == 0:
        outputPdf = f'{name}_sel_{numPgOne + 1}_{numPgTwo}{ext}'
        listSel = [pg for pg in range(numPgOne, numPgTwo)]
    else:
        numPages = inputPdf.page_count
        outputPdf = f'{name}_del_{numPgOne + 1}_{numPgTwo}{ext}'
        listSel = [pg for pg in range(numPages) if pg not in range(numPgOne, numPgTwo)]
    docPdf.select(listSel)
    docPdf.save(outputPdf)
    docPdf.close()
    if index != 4:
        outputPdf = rotatePdf(outputPdf, index)       
    downPdfUnique(outputPdf, numPgOne, numPgTwo, namePdf) 
                       
def downPdfUnique(outputPdf, numPgOne, numPgTwo, namePdf):
    with open(outputPdf, "rb") as pdf_file:
        PDFbyte = pdf_file.read()
    if len(PDFbyte) > 0:
        mensResult(0, 1, 'pdf', PDFbyte, outputPdf)
        
def extractPgs(docPdf, numPgOne, numPgTwo, mode, namePdf, index):
    numPgOne -= 1
    filesPdf = [docPdf]
    filesRead = []  
    name, ext = os.path.splitext(namePdf)
    for file in filesPdf:
        numPages = file.page_count
        diffPg = abs(numPgTwo - numPgOne)
        minPg = min([numPgOne, numPgTwo])
        listPg = [pg for pg in range(minPg, diffPg)]
        for p, pageNum in enumerate(listPg):
            inputPdf = file
            outputPdf = f'{name}_{pageNum + 1}.pdf'
            newPdf = pymupdf.open()
            newPdf.insert_pdf(inputPdf, from_page=pageNum, to_page=pageNum)
            newPdf.save(outputPdf)
            if index != 4:
                outputPdf = rotatePdf(outputPdf, index) 
            filesRead.append(outputPdf)
            newPdf.close()
    downloadExt(filesRead, namePdf, numPgOne, numPgTwo, 'páginas')

def exibeInfo(docPdf):
    @st.dialog(' ')
    def config():
        trace= '_'*10
        nPgs = docPdf.page_count
        size = round(uploadPdf.size/1024, 2)
        if size > 1024:
            size /= 1024
            size = round(size, 2)
            unit = 'MB'
        else:
            unit = 'KB'
        typFile = uploadPdf.type
        dirty = docPdf.is_dirty
        pdfYes = docPdf.is_pdf
        close = docPdf.is_closed
        formPdf = docPdf.is_form_pdf
        encry = docPdf.is_encrypted
        pdfMeta = docPdf.metadata
        st.markdown(f'🗄️ **Tamanho**: {size}{unit}') 
        st.markdown(f'📄️ **Total de páginas**: {nPgs}')
        dictKeys = {'creator': '💂 **criador**', 'producer': '🔴 **responsável**', 'creationDate': '📅 **dia de criação**', 
                    'modDate': '🕰️ **dia de modificação**', 'title': '#️⃣  **título**', 'author': '📕 **autor**', 'format': '⏹️ **formato**',
                    'subject': '🖊️ **assunto**', 'keywords': '#️⃣  **palavras-chave**', 'encryption': '🔑 **criptografia**'}
        keys = [key for key in list(dictKeys.keys())]
        for key in keys:
            valueKey = dictKeys[key]
            metaKey = pdfMeta[key]
            if metaKey is None:
                metaKey = trace
            else:
                if len(metaKey.strip()) == 0:
                    metaKey = trace
            st.markdown(f'{dictKeys[key]}: {metaKey}')
            
        #st.write(os.path.getsize(filePdf))
        #st.markdown(f'🔴 **Número de acessos a módulos/submódulos do app**: {info[infoKeys[0]]}')
        #st.markdown(f'📅 **Data de início da sessão**: {dateFullLang(info[infoKeys[1]])}')
        #st.markdown(f'⏳ **Tempo de uso deste app**: {tempus} segundo(s).')
        #st.markdown(f"📐 **Clique(s) no botão :blue[Cálculo]**: {calcK}.")
        #st.markdown(f"👉 **Clique(s) no botão :blue[Arquivos]**: {fileK}.")
        #st.markdown(f"👓 **Clique(s) no botão :blue[Acesso]**: {accessK}.")
        #st.markdown(f"#️⃣ **Clique(s) no botão :blue[Feriados]**: {holidayK}.")
        #st.markdown(f"📕 **Clique(s) no botão :blue[Limpeza]**: {clearK}.")
    config()
                
@st.dialog(' ')
def config(str):
    st.text(str)           
        
def iniFinally(mode):
    if mode == 0:
        for key in listKeys:
            if key not in st.session_state:
                try:
                    st.session_state[key] = dictKeys[key]
                except:
                    pass        
    else:
        try:
            for key in listKeys:
                del st.session_state[key]
        except:
            pass  
        iniFinally(0)
        st.rerun()

def main():
    global uploadPdf
    global valMx
    global sufix
    sufix = ['']
    with st.container(border=6):
        uploadPdf = st.file_uploader('Selecionar arquivos PDF', 
                                     type=['pdf'], 
                                     accept_multiple_files=False)
        if uploadPdf is not None:
            pdfName = uploadPdf.name
            docPdf = pymupdf.open(stream=uploadPdf.read(), filetype="pdf")
            valMx = docPdf.page_count 
            valMxSize = round(uploadPdf.size/(1024**2), 2)
            colPgOne, colPgTwo, colSlider, colSize, colMark = st.columns([1.5, 1.5, 2.3, 1.5, 3.2])
            numPgOne = colPgOne.number_input(label='Página inicial', key=listKeys[0], 
                                             min_value=1, max_value=valMx)
            numPgTwo = colPgTwo.number_input(label='Página final', key=listKeys[1], 
                                             min_value=1, max_value=valMx)
            valPgAngle = colSlider.select_slider(label='Ângulo de rotação', options=valAngles, 
                                                 key=listKeys[2])    
            valPgSize = colSize.number_input(label='Tamanho para divisão (Mb)', key=listKeys[3], 
                                             min_value=dictKeys[listKeys[3]], step=dictKeys[listKeys[3]],  
                                             max_value=valMxSize)
            valPgMark = colMark.text_input(label="Marca d'água", key=listKeys[4], max_chars=50, 
                                           value=dictKeys[listKeys[4]])
            colButtAct, colButtTxt, colButtSel, colButtDel, colButtClear = st.columns(5)
            buttPgAct = colButtAct.button(label='Corte/páginas', key=keysButts[0], 
                                          use_container_width=True, icon=":material/cut:")
            buttPgTxt = colButtTxt.button(label='Texto', key=keysButts[1], 
                                          use_container_width=True, icon=":material/description:")
            buttPgSel = colButtSel.button(label='Seleção', key=keysButts[2], 
                                          use_container_width=True, icon=":material/list:")
            buttPgDel = colButtDel.button(label='Deleção', key=keysButts[3], 
                                          use_container_width=True, icon=":material/delete:")
            buttPgClear = colButtClear.button(label='Limpeza', key=keysButts[4], 
                                              use_container_width=True, icon=":material/square:")
            colButtUrl, colButtImg, colButtSize, colButtMark, colButtInfo = st.columns(5)
            buttPdfUrl = colButtUrl.button(label='URLs', key=keysButts[5], 
                                           use_container_width=True, icon=":material/link:")
            buttPdfImg = colButtImg.button(label='Imagens', key=keysButts[6], 
                                           use_container_width=True, icon=":material/image:")
            buttPdfSize = colButtSize.button(label='Corte/tamanho', key=keysButts[7], 
                                           use_container_width=True, icon=":material/docs:")
            buttPdfMark = colButtMark.button(label='Marcação', key=keysButts[8], 
                                             use_container_width=True, icon=":material/approval:")
            buttPdfInfo =  colButtInfo.button(label='Informações', key=keysButts[9], 
                                           use_container_width=True, icon=":material/info:")
            if numPgTwo >= numPgOne: 
                numPgIni = numPgOne
                numPgFinal = numPgTwo
            else:
                numPgIni = numPgTwo
                numPgFinal = numPgOne 
            indexAng = valAngles.index(valPgAngle)
            exprPre = f'o intervalo de páginas {numPgOne} a {numPgTwo} do arquivo \n{pdfName}'
            if buttPgAct:  
                with st.spinner(f'Dividindo {exprPre}!'):
                    extractPgs(docPdf, numPgIni, numPgFinal, 0, pdfName, indexAng)
            if buttPgTxt: 
                with st.spinner(f'Extraindo texto d{exprPre}!'):
                    selTxtUrlPgs(docPdf, numPgOne, numPgTwo, pdfName, 0, indexAng)
            if buttPdfUrl:
                with st.spinner(f'Extraindo links/URLs d{exprPre}!'):
                    sufix[0] = 'urls'
                    selTxtUrlPgs(docPdf, numPgOne, numPgTwo, pdfName, 1, indexAng)
            if buttPdfImg: 
                with st.spinner(f'Extraindo imagens d{exprPre}!'):
                    sufix[0] = 'imgs'
                    selImgUrlsPgs(docPdf, numPgOne, numPgTwo, pdfName, 2, indexAng)
            if buttPgSel:
                with st.spinner(f'Criando arquivo com {exprPre}!'):
                    selDelPgs(docPdf, numPgOne, numPgTwo, pdfName, 0, indexAng)
            if buttPgDel: 
                with st.spinner(f'Criando arquivo com deleção d{exprPre}!'):
                    selDelPgs(docPdf, numPgOne, numPgTwo, pdfName, 1, indexAng)
            if buttPdfSize:
                with st.spinner(f'Dividindo {exprPre} em pedaços de {valPgSize}Mb!'):
                    selPgsSize(docPdf, numPgOne, numPgTwo, pdfName, indexAng, valPgSize)
            if buttPdfMark:
                if valPgMark.strip() == '':
                    valPgMark = nameApp 
                with st.spinner(f"Carimbando {exprPre} com a marca d'água {valPgMark}."):
                    selPdfMark(docPdf, numPgOne, numPgTwo, pdfName, indexAng, valPgMark)
            if buttPdfInfo:
                exibeInfo(docPdf)
            if buttPgClear:                
                iniFinally(1)  
        
if __name__ == '__main__':
    global dictKeys, listKeys 
    global keysButts, valAngles, valComps
    global countPg
    global namesTeste, nameApp 
    nameApp = 'Ferramentas/PDF'
    valAngles = ['-360°', '-270°', '-180°', '-90°', '0°', '90°', '180°', '270°', '360°']
    #valComps = ['mínimo', 'regular', 'bom', 'muito bom', 'ótimo', 'radical']    
    dictKeys = {'pgOne': 1, 
                'pgTwo': 1, 
                'pgAngle': valAngles[0], 
                'pgSize': 0.05, 
                'pgMark': ''}
    listKeys = list(dictKeys.keys())
    keysButts = ['buttAct', 'buttTxt', 'buttSel', 'buttDel', 'buttClear', 
                 'buttUrls', 'buttImgs', 'buttSize', 'buttCompress', 'buttOCR']
    countPg = []
    namesTeste = []
    dirBin = r'C:\Users\ACER\Documents\bin'
    st.set_page_config(page_title=nameApp,  page_icon=":material/files:", 
                       layout='wide')
    st.cache_data.clear() 
    iniFinally(0)
    main()
