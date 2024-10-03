package com.viacerta.service;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.FileNotFoundException;

import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class LibreOfficeConverterService {
    public static File convertToPDF(XWPFDocument document) throws Exception {
        // Obter o diretório do projeto
        String projectDir = System.getProperty("user.dir");
        
        // Criar arquivo temporário .docx na raiz do projeto
        File tempDocFile = new File(projectDir, "temp_" + System.currentTimeMillis() + ".docx");
        
        try (FileOutputStream out = new FileOutputStream(tempDocFile)) {
            document.write(out);
        }

        // Comando para converter usando LibreOffice
        String[] command = {"libreoffice", "--headless", "--convert-to", "pdf", tempDocFile.getAbsolutePath()};
        Process process = new ProcessBuilder(command)
            .directory(new File(projectDir)) // Definir diretório de trabalho
            .start();
        process.waitFor();

        // Aguardar um pouco para garantir que o sistema de arquivos sincronize
        //Thread.sleep(2000);

        // Procurar o arquivo PDF na raiz do projeto
        String pdfFileName = tempDocFile.getName().replace(".docx", ".pdf");
        File pdfFile = new File(projectDir, pdfFileName);

        if (!pdfFile.exists()) {
            File[] possiblePdfs = new File(projectDir).listFiles((dir, name) -> name.startsWith("temp_") && name.endsWith(".pdf"));
            
            if (possiblePdfs != null && possiblePdfs.length > 0) {
                pdfFile = possiblePdfs[0];
            } else {
                throw new FileNotFoundException("Erro na conversão para PDF. Arquivo não encontrado na raiz do projeto.");
            }
        }

        deleteFile(tempDocFile);

        return pdfFile;
    }

    public static byte[] unifyDocuments(File mainPdf, File subReportPdf) throws IOException {
        PDFMergerUtility pdfMerger = new PDFMergerUtility();
        pdfMerger.addSource(new FileInputStream(mainPdf));
        if (subReportPdf != null && subReportPdf.exists()) { 
            pdfMerger.addSource(new FileInputStream(subReportPdf));
        }
        
        ByteArrayOutputStream mergedPdfOutputStream = new ByteArrayOutputStream();
        pdfMerger.setDestinationStream(mergedPdfOutputStream);
        pdfMerger.mergeDocuments(MemoryUsageSetting.setupMainMemoryOnly());
        
        return mergedPdfOutputStream.toByteArray();
    }

    public static void deleteFile(File file) {
        if (file != null && file.exists()) {
            if (!file.delete()) {
                System.err.println("Não foi possível deletar o arquivo temporário: " + file.getAbsolutePath());
                file.deleteOnExit();
            }
        }
    }
}
