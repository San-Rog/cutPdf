import pymupdf
import streamlit as st
import zipfile
import os
import datetime
from PyPDF2 import PdfReader, PdfWriter

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

def downloadExt(files):
    fileTmp = f'{nameFile()}_tempFile.zip'
    fileZip = f'file_{nameFile()}.zip'
    for file in files:
        with open(file, "rb") as extFile:
           PDFbyte = extFile.read()
        with zipfile.ZipFile(fileTmp, 'a') as extFile:
           extFile.writestr(file, PDFbyte)
    nFiles = len(fileTmp)
    if nFiles > 0:
        st.success(f'Gerado(s) {len(files)} arquivo(s)', icon='ℹ️')
        with open(fileTmp, "rb") as file:
            st.download_button(label="Download_zip",
                               data=file,
                               file_name=fileZip,
                               mime='application/zip', 
                               icon=":material/download:")

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
    
def savePdf(outputBase, contPartes, writer):
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
                outputPdf = savePdf(outputBase, contPartes, writer)
                filesCutSave.append(outputPdf)
                writer = PdfWriter()
                sizeActual = 0
                contPartes += 1
        if len(writer.pages) > 0:
            outputPdf = savePdf(outputBase, contPartes, writer)
            filesCutSave.append(outputPdf)
    except Exception as e:
        st.error(f"Ocorreu um erro: {e} - página {i+1}", icon='🛑')
    return filesCutSave    

def createPdfSel(docPdf, numPgOne, numPgTwo, namePdf, index):
    numPgOne -= 1    
    inputPdf = docPdf
    name, ext = os.path.splitext(namePdf)
    outputPdf = 'temp_sel.pdf'
    listSel = [pg for pg in range(numPgOne, numPgTwo)]
    docPdf.select(listSel)
    docPdf.save(outputPdf)
    if index != 4:
        outputPdf = rotatePdf(outputPdf, index) 
    docPdf.close()
    return outputPdf    
    
def selPgsSize(docPdf, numPgOne, numPgTwo, namePdf, index, sizeMax):
    outputPdf = createPdfSel(docPdf, numPgOne, numPgTwo, namePdf, index)
    inputPdf = outputPdf
    sizeMaxStr = str(sizeMax).replace('.', '_')
    sizeTeste = str(sizeMax).split('.')
    st.write(sizeTeste)
    outputBase = f'{os.path.splitext(inputPdf)[0]}_divisão_{sizeMaxStr}_Mb__parte_'
    filesCutSave = divideBySize(inputPdf, sizeMax, outputBase)
    downloadExt(filesCutSave)
        
def selImgUrlsPgs(docPdf, numPgOne, numPgTwo, namePdf, mode, index):
    outputPdf = createPdfSel(docPdf, numPgOne, numPgTwo, namePdf, index)
    filesImg = extractImgs(outputPdf)
    downloadExt(filesImg)
    
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
        st.download_button(label=strLabel,
                           data=text,
                           file_name=outputTxt,
                           mime="text/csv", 
                           icon=":material/download:")
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
    with open(outputPdf, "rb") as pdf_file:
        PDFbyte = pdf_file.read()
    st.download_button(label='Download_pdf', 
                       data=PDFbyte,
                       file_name=outputPdf,
                       mime='application/octet-stream', 
                       icon=":material/download:") 

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
    downloadExt(filesRead)
                
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
    with st.container(border=6):
        uploadPdf = st.file_uploader('Selecionar arquivos PDF', 
                                     type=['pdf'], 
                                     accept_multiple_files=False)
        if uploadPdf is not None:
            pdfName = uploadPdf.name
            docPdf = pymupdf.open(stream=uploadPdf.read(), filetype="pdf")
            valMx = docPdf.page_count 
            valMxSize = round(uploadPdf.size/(1024**2), 2)
            colPgOne, colPgTwo, colSlider, colSize = st.columns(4)
            numPgOne = colPgOne.number_input(label='Página inicial', key=listKeys[0], 
                                             min_value=1, max_value=valMx)
            numPgTwo = colPgTwo.number_input(label='Página final', key=listKeys[1], 
                                             min_value=1, max_value=valMx)
            valPgAngle = colSlider.select_slider(label='Ângulo de rotação', options=valAngles, 
                                                 key=listKeys[2], value=dictKeys[listKeys[2]])     
            valPgSize = colSize.number_input(label='Tamanho para divisão (Mb)', key=listKeys[3], 
                                             value=0.05, step=0.05, max_value=valMxSize)
            colButtAct, colButtTxt, colButtSel, colButtDel, colButtClear = st.columns(5)
            buttPgAct = colButtAct.button(label='Corte/páginas', key=keysButts[0], use_container_width=True, icon=":material/cut:")
            buttPgTxt = colButtTxt.button(label='Texto', key=keysButts[1], use_container_width=True, icon=":material/description:")
            buttPgSel = colButtSel.button(label='Seleção', key=keysButts[2], use_container_width=True, icon=":material/list:")
            buttPgDel = colButtDel.button(label='Deleção', key=keysButts[3], use_container_width=True, icon=":material/delete:")
            buttPgClear = colButtClear.button(label='Limpeza', key=keysButts[4], use_container_width=True, icon=":material/square:")
            colButtUrl, colButtImg, colButtSize, colButtB, colButtC = st.columns(5)
            buttPdfUrl = colButtUrl.button(label='URLs', key=keysButts[5], use_container_width=True, icon=":material/link:")
            buttPdfImg = colButtImg.button(label='Imagens', key=keysButts[6], use_container_width=True, icon=":material/image:")
            buttPdfSize = colButtSize.button(label='Corte/tamanho', key=keysButts[7], use_container_width=True, icon=":material/docs:")
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
                    selTxtUrlPgs(docPdf, numPgOne, numPgTwo, pdfName, 1, indexAng)
            if buttPdfImg: 
                with st.spinner(f'Extraindo imagens d{exprPre}!'):
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
            if buttPgClear:
                iniFinally(1)                
        
if __name__ == '__main__':
    global dictKeys, listKeys 
    global keysButts, valAngles
    global countPg
    global namesTeste
    valAngles = ['-360°', '-270°', '-180°', '-90°', '0°', '90°', '180°', '270°', '360°']
    dictKeys = {'pgOne': 1, 
                'pgTwo': 1, 
                'pgAngle': valAngles[4], 
                'pgSize': 0.05}
    listKeys = list(dictKeys.keys())
    keysButts = ['buttAct', 'buttTxt', 'buttSel', 'buttDel', 'buttClear', 
                 'buttUrls', 'buttImgs', 'buttSize']
    countPg = []
    namesTeste = []
    st.set_page_config(page_title='Ferramentas de tratamento de PDF',  page_icon=":material/files:", 
                       layout='wide')
    st.cache_data.clear() 
    iniFinally(0)
    main()
