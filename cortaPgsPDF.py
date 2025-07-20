import pymupdf
import streamlit as st
import zipfile
import os
import datetime

@st.cache_data
def extractText(filePdf, text):
    docPdf = pymupdf.open(filePdf)
    for page in docPdf:
        text += page.get_text()
    docPdf.close()
    return text

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

def downloadPdf(filesPdf):
    fileTmp = f'{nameFile()}_tempFile.zip'
    fileZip = f'file_{nameFile()}.zip'
    for filePdf in filesPdf:
        with open(filePdf, "rb") as pdf_file:
            PDFbyte = pdf_file.read()
        with zipfile.ZipFile(fileTmp, 'a') as pdf_file:
            pdf_file.writestr(filePdf, PDFbyte)
    nFiles = len(fileTmp)
    if nFiles > 0:
        st.success(f'Gerado(s) {len(filesPdf)} arquivo(s)', icon='ℹ️')
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
                               
def selDelPgs(docPdf, numPgOne, numPgTwo, namePdf, mode, index):
    numPgOne -= 1
    inputPdf = docPdf
    name, ext = os.path.splitext(namePdf)
    if mode == 0:
        outputPdf = f'{name}_sel_{numPgOne}_{numPgTwo}{ext}'
        listSel = [pg for pg in range(numPgOne, numPgTwo)]
    else:
        numPages = inputPdf.page_count
        outputPdf = f'{name}_del_{numPgOne}_{numPgTwo}{ext}'
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
    if mode == 1:
        text = ''
    for file in filesPdf:
        numPages = file.page_count
        diffPg = abs(numPgTwo - numPgOne)
        minPg = min([numPgOne, numPgTwo])
        listPg = [pg for pg in range(minPg, diffPg)]
        for pageNum in listPg:
            inputPdf = file
            outputPdf = f'{name}_{pageNum + 1}.pdf'
            newPdf = pymupdf.open()
            newPdf.insert_pdf(inputPdf, from_page=pageNum, to_page=pageNum)
            newPdf.save(outputPdf)
            if mode == 1:
                text += extractText(outputPdf, text)                
            newPdf.close()
            if index != 4:
                outputPdf = rotatePdf(outputPdf, index) 
            filesRead.append(outputPdf)
    if mode == 0:
        downloadPdf(filesRead)
    else:
        outputPdf = f'{name}_{numPgOne + 1}_{numPgTwo}.txt'
        st.download_button(label="Download_txt",
                           data=text,
                           file_name=outputPdf,
                           mime="text/csv", 
                           icon=":material/download:")
    
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
    uploadPdf = st.file_uploader('Selecionar arquivos PDF', 
                                 type=['pdf'], 
                                 accept_multiple_files=False)
    if uploadPdf is not None:
        pdfName = uploadPdf.name
        docPdf = pymupdf.open(stream=uploadPdf.read(), filetype="pdf")
        valMx = docPdf.page_count  
        colPgOne, colPgTwo, colSlider = st.columns(3)
        numPgOne = colPgOne.number_input(label='Página inicial', key=listKeys[0], 
                                         min_value=1, max_value=valMx)
        numPgTwo = colPgTwo.number_input(label='Página final', key=listKeys[1], 
                                         min_value=1, max_value=valMx)
        valPgAngle = colSlider.select_slider(label='Ângulo de rotação', options=valAngles, key=listKeys[2], value='0°')
        colButtAct, colButtTxt, colButtSel, colButtDel, colClear = st.columns(5)
        buttPgAct = colButtAct.button(label='Divisão', key=keysButts[0], use_container_width=True, icon=":material/cut:")
        buttPgTxt = colButtTxt.button(label='Texto', key=keysButts[1], use_container_width=True, icon=":material/description:")
        buttPgSel = colButtSel.button(label='Seleção', key=keysButts[2], use_container_width=True, icon=":material/list:")
        buttPgDel = colButtDel.button(label='Deleção', key=keysButts[3], use_container_width=True, icon=":material/delete:")
        buttPgClear = colClear.button(label='Limpeza', key=keysButts[4], use_container_width=True, icon=":material/square:")
        if numPgTwo >= numPgOne: 
            numPgIni = numPgOne
            numPgFinal = numPgTwo
        else:
            numPgIni = numPgTwo
            numPgFinal = numPgOne 
        indexAng = valAngles.index(valPgAngle)
        if buttPgAct:   
            extractPgs(docPdf, numPgIni, numPgFinal, 0, pdfName, indexAng)
        if buttPgTxt: 
            extractPgs(docPdf, numPgIni, numPgFinal, 1, pdfName, indexAng)
        if buttPgSel:
            selDelPgs(docPdf, numPgOne, numPgTwo, pdfName, 0, indexAng)
        if buttPgDel: 
            selDelPgs(docPdf, numPgOne, numPgTwo, pdfName, 1, indexAng)
        if buttPgClear:
            iniFinally(1)        
    
if __name__ == '__main__':
    global dictKeys, listKeys 
    global keysButts, valAngles
    dictKeys = {'pgOne': 1, 
                'pgTwo': 1, 
                'pgAngle': '0°'}
    listKeys = list(dictKeys.keys())
    keysButts = ['buttAct', 'buttTxt', 'buttSel', 'buttDel', 'buttClear']
    valAngles = ['-360°', '-270°', '-180°', '-90°', '0°', '90°', '180°', '270°', '360°']
    st.set_page_config(page_title='Ferramentas de tratamento de PDF',  page_icon=":material/files:")
    st.cache_data.clear() 
    iniFinally(0)
    main()



