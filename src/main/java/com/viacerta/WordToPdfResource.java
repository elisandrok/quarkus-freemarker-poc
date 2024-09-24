package com.viacerta;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.jboss.resteasy.reactive.MultipartForm;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.WebColors;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Text;
import com.itextpdf.layout.properties.TextAlignment;

import freemarker.template.Configuration;
import freemarker.template.Template;
import jakarta.ws.rs.Consumes;
import jakarta.ws.rs.POST;
import jakarta.ws.rs.Path;
import jakarta.ws.rs.Produces;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.Response;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.StringWriter;
import java.util.Map;

@Path("/convert")
public class WordToPdfResource {

    private static final Logger sl4jLogger = LoggerFactory.getLogger(WordToPdfResource.class);


    @POST
    @Consumes(MediaType.MULTIPART_FORM_DATA)
    @Produces("application/pdf")
    public Response convertWordToPdf(@MultipartForm WordDataForm form) {
        try {
            // Process the uploaded file and substitute variables
            InputStream docxInputStream = form.getFile();
            Map<String, Object> variables = form.getVariables();

            byte[] pdfContent = processWordTemplate(docxInputStream, variables);

            // Return the generated PDF
            return Response.ok(pdfContent)
                    .header("Content-Disposition", "attachment; filename=converted.pdf")
                    .build();
        } catch (Exception e) {
            e.printStackTrace();
            return Response.serverError().entity("Error processing the file").build();
        }
    }

    private byte[] processWordTemplate(InputStream docxInputStream, Map<String, Object> variables) throws Exception {
        // Carrega o documento .docx
        XWPFDocument document = new XWPFDocument(docxInputStream);

        // Processa cabeçalhos e rodapés
        processHeadersAndFooters(document, variables);
        
        // Iterar sobre os parágrafos e aplicar o template Freemarker
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            String text = paragraph.getText();
            String processedText = processTemplate(text, variables); // Processa o texto com Freemarker
            
            for (XWPFRun run : paragraph.getRuns()) {
                run.setText(processedText, 0);
            }
        }
        
        // Converter para PDF após substituir variáveis
        ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream();
        convertToPdf(document, pdfOutputStream);

        return pdfOutputStream.toByteArray();
    }

    private void convertToPdf(XWPFDocument document, ByteArrayOutputStream pdfOutputStream) throws Exception {
        // Inicia o escritor de PDF
        PdfWriter writer = new PdfWriter(pdfOutputStream);
        PdfDocument pdfDoc = new PdfDocument(writer);
        Document pdfDocument = new Document(pdfDoc);

        // Itera sobre os parágrafos do documento Word e escreve no PDF
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            System.out.println("Paragrafo: " + paragraph.getText());
            Paragraph pdfParagraph = new Paragraph();

            // Configurar alinhamento
            pdfParagraph.setTextAlignment(getPdfTextAlignment(paragraph.getAlignment()));

            // Adicionar estilos e runs
            for (XWPFRun run : paragraph.getRuns()) {
                Text text = new Text(run.getText(0));
                // Configurar cor
                if (run.getColor() != null) {
                    text.setFontColor(WebColors.getRGBColor(run.getColor())); // Use a cor especificada
                }
                // Configurar negrito, itálico, sublinhado
                if (run.isBold()) {
                    text.setBold();
                }
                if (run.isItalic()) {
                    text.setItalic();
                }
                if (run.getUnderline() != UnderlinePatterns.NONE) {
                    text.setUnderline();
                }

                pdfParagraph.add(text);
            }            
            pdfDocument.add(pdfParagraph);
        }

        // Processar tabelas se houver
        for (XWPFTable table : document.getTables()) {
            convertTableToPdf(table, pdfDocument);
        }

        processImages(document, pdfDocument);
        // Finaliza o documento PDF
        pdfDocument.close();
    }
    
    private String processTemplate(String templateContent, Map<String, Object> variables) throws Exception {
        Configuration cfg = new Configuration(Configuration.VERSION_2_3_31);
        Template template = new Template("docxTemplate", templateContent, cfg);

        StringWriter writer = new StringWriter();
        template.process(variables, writer);

        return writer.toString();
    }

    private void processHeadersAndFooters(XWPFDocument document, Map<String, Object> variables) throws Exception {
        // Processa cabeçalhos
        for (XWPFHeader header : document.getHeaderList()) {
            for (XWPFParagraph paragraph : header.getParagraphs()) {
                String text = paragraph.getText();
                String processedText = processTemplate(text, variables);
                for (XWPFRun run : paragraph.getRuns()) {
                    run.setText(processedText, 0);
                }
            }
        }

        // Processa rodapés
        for (XWPFFooter footer : document.getFooterList()) {
            for (XWPFParagraph paragraph : footer.getParagraphs()) {
                String text = paragraph.getText();
                String processedText = processTemplate(text, variables);
                for (XWPFRun run : paragraph.getRuns()) {
                    run.setText(processedText, 0);
                }
            }
        }
    }

    private void processImages(XWPFDocument document, Document pdfDocument) throws Exception {
        for (XWPFPictureData pictureData : document.getAllPictures()) {
            // String imgFileName = pictureData.suggestFileExtension();
            byte[] imageBytes = pictureData.getPackagePart().getInputStream().readAllBytes();
            
            ImageData imageData = ImageDataFactory.create(imageBytes);
            Image pdfImage = new Image(imageData);

            // Você pode ajustar a largura e altura da imagem se necessário
            // pdfImage.scaleToFit(500, 500); // ajuste conforme necessário
            pdfDocument.add(pdfImage);
        }
    }

    private void convertTableToPdf(XWPFTable table, Document pdfDocument) {
        int numColumns = table.getRow(0).getTableCells().size(); // Obtém o número de colunas da primeira linha
        Table pdfTable = new Table(numColumns); // Cria a tabela com o número de colunas
    
        // Itera pelas linhas da tabela
        for (XWPFTableRow row : table.getRows()) {
            // Itera pelas células de cada linha
            for (XWPFTableCell cell : row.getTableCells()) {
                Cell pdfCell = new Cell(); // Cria uma nova célula para o PDF
                
                // Adiciona o texto da célula ao pdfCell
                pdfCell.add(new Paragraph(cell.getText()));
                
                // Adiciona a célula à tabela PDF
                pdfTable.addCell(pdfCell);
            }
        }
    
        // Adiciona a tabela PDF ao documento
        pdfDocument.add(pdfTable);
    }


    private TextAlignment getPdfTextAlignment(ParagraphAlignment alignment) {
        switch (alignment) {
            case CENTER:
                return TextAlignment.CENTER;
            case RIGHT:
                return TextAlignment.RIGHT;
            case LEFT:
            default:
                return TextAlignment.LEFT;
        }
    }


}
