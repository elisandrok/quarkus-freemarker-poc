package com.viacerta.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileNotFoundException;
import java.util.Date;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class LibreOfficeConverterService {
    public static File convertToPDF(XWPFDocument document) throws Exception {
        System.out.println("Iniciando conversão para PDF");
        
        // Obter o diretório do projeto
        String projectDir = System.getProperty("user.dir");
        System.out.println("Diretório do projeto: " + projectDir);
        
        // Criar arquivo temporário .docx na raiz do projeto
        File tempDocFile = new File(projectDir, "temp_" + System.currentTimeMillis() + ".docx");
        System.out.println("Arquivo temporário criado: " + tempDocFile.getAbsolutePath());
        
        try (FileOutputStream out = new FileOutputStream(tempDocFile)) {
            document.write(out);
            System.out.println("Documento escrito no arquivo temporário");
        }

        // Comando para converter usando LibreOffice
        String[] command = {"libreoffice", "--headless", "--convert-to", "pdf", tempDocFile.getAbsolutePath()};
        System.out.println("Comando de conversão: " + String.join(" ", command));
        
        Process process = new ProcessBuilder(command)
            .directory(new File(projectDir)) // Definir diretório de trabalho
            .start();
        int exitCode = process.waitFor();
        System.out.println("Processo de conversão concluído com código de saída: " + exitCode);

        // Aguardar um pouco para garantir que o sistema de arquivos sincronize
        Thread.sleep(2000);

        // Procurar o arquivo PDF na raiz do projeto
        String pdfFileName = tempDocFile.getName().replace(".docx", ".pdf");
        File pdfFile = new File(projectDir, pdfFileName);
        System.out.println("Procurando arquivo PDF em: " + pdfFile.getAbsolutePath());

        if (!pdfFile.exists()) {
            System.out.println("PDF não encontrado, procurando por variações do nome...");
            File[] possiblePdfs = new File(projectDir).listFiles((dir, name) -> name.startsWith("temp_") && name.endsWith(".pdf"));
            
            if (possiblePdfs != null && possiblePdfs.length > 0) {
                pdfFile = possiblePdfs[0];
                System.out.println("PDF encontrado com nome alternativo: " + pdfFile.getAbsolutePath());
            } else {
                System.err.println("Arquivo PDF não encontrado após todas as tentativas");
                System.out.println("Conteúdo do diretório do projeto:");
                File[] files = new File(projectDir).listFiles();
                if (files != null) {
                    for (File file : files) {
                        System.out.println(file.getName() + " - Existe: " + file.exists() + ", Pode ler: " + file.canRead());
                    }
                } else {
                    System.out.println("Não foi possível listar os arquivos do diretório");
                }
                throw new FileNotFoundException("Erro na conversão para PDF. Arquivo não encontrado na raiz do projeto.");
            }
        }

        System.out.println("Arquivo PDF encontrado: " + pdfFile.getAbsolutePath());
        System.out.println("Tamanho do arquivo: " + pdfFile.length() + " bytes");
        System.out.println("Última modificação: " + new Date(pdfFile.lastModified()));

        // Limpeza: remover o arquivo .docx temporário
        if (!tempDocFile.delete()) {
            System.out.println("Não foi possível deletar o arquivo temporário .docx");
        }

        return pdfFile;
    }
}
