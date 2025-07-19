import pymupdf  
import streamlit as st
import zipfile
import os
import datetime

def extractText(docPdf):
    text = ''
    for page in docPdf:
        text += page.get_text()
    st.text_area(label='Transcrição do PDF', value=text, height=400) 
    docPdf.close()
    
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
    fileZip = "mydownload.zip"
    for filePdf in filesPdf:
        with open(filePdf, "rb") as pdf_file:
            PDFbyte = pdf_file.read()
        with zipfile.ZipFile(fileTmp, 'a') as pdf_file:
            pdf_file.writestr(filePdf, PDFbyte)
    with open(fileTmp, "rb") as file:
        st.download_button(label="Download_ZIP",
                       data=file,
                       file_name=fileZip,
                       mime='application/zip')

def extractPgs(docPdf):
    filesPdf = [docPdf]
    filesRead = []
    for file in filesPdf:
        numPages = file.page_count
        for pageNum in range(numPages):
            inputPdf = file
            outputPdf = f'arquivo_{pageNum}.pdf'
            new_pdf = pymupdf.open()
            new_pdf.insert_pdf(inputPdf, from_page=pageNum, to_page=pageNum)
            new_pdf.save(outputPdf)
            new_pdf.close()
            st.text(f'Divisão da página {pageNum + 1} salva em {outputPdf}')
            #downloadPdf(filePdf)
            filesRead.append(outputPdf)
    downloadPdf(filesRead)
    
        
def main():
    uploadPdf = st.file_uploader('Selecionar arquivos PDF', type=['pdf'], accept_multiple_files=False)
    if uploadPdf is not None: 
        docPdf = pymupdf.open(stream=uploadPdf.read(), filetype="pdf")
        #extractText(docPdf)
        extractPgs(docPdf)

if __name__ == '__main__':
    main()



