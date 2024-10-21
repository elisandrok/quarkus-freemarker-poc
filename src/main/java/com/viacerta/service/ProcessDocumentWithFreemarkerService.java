package com.viacerta.service;

import java.io.IOException;
import java.io.StringWriter;
import java.util.Map;
import java.util.TimeZone;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import freemarker.cache.StringTemplateLoader;
import freemarker.core.ParseException;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import freemarker.template.TemplateExceptionHandler;
import jakarta.enterprise.context.ApplicationScoped;

@ApplicationScoped
public class ProcessDocumentWithFreemarkerService {
    public void processDocumentWithFreemarker(XWPFDocument document, Map<String, Object> fieldsJson)
            throws IOException, TemplateException {
        Configuration cfg = new Configuration(Configuration.VERSION_2_3_33);
        cfg.setDefaultEncoding("UTF-8");
        // cfg.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
        cfg.setTemplateExceptionHandler(TemplateExceptionHandler.HTML_DEBUG_HANDLER);// for development
        cfg.setSQLDateAndTimeTimeZone(TimeZone.getDefault());

        // processar cabeçalhos
        for (XWPFHeader header : document.getHeaderList()) {
            processHeaderWithImages(header, cfg, fieldsJson);
        }

        // Processar parágrafos
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            processParagraphWithFreemarker(paragraph, cfg, fieldsJson);
        }

        // Processar tabelas
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        processParagraphWithFreemarker(paragraph, cfg, fieldsJson);
                    }
                }
            }
        }
    }

    private void processHeaderWithImages(XWPFHeader header, Configuration cfg, Map<String, Object> fieldsJson)
            throws IOException, TemplateException {
        for (XWPFParagraph paragraph : header.getParagraphs()) {
            StringBuilder fullText = new StringBuilder();
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                if (run.getEmbeddedPictures().isEmpty()) {
                    String text = run.getText(0);
                    if (text != null) {
                        fullText.append(text);
                    }
                }
            }

            // Processar o texto completo
            String processedText = processTextWithFreemarker(fullText.toString(), cfg, fieldsJson);

            // Se o texto foi modificado, atualizar os runs
            if (!fullText.toString().equals(processedText)) {
                // Limpar o conteúdo existente
                for (int i = runs.size() - 1; i >= 0; i--) {
                    paragraph.removeRun(i);
                }

                // Adicionar o novo texto processado
                XWPFRun newRun = paragraph.createRun();
                newRun.setText(processedText);
            }
        }
    }

    private String processTextWithFreemarker(String text, Configuration cfg, Map<String, Object> fieldsJson)
            throws IOException, TemplateException {
        if (text == null || text.trim().isEmpty()) {
            return text;
        }

        // Verifica se há placeholders FreeMarker completos no texto
        if (!text.matches(".*\\$\\{\\w+\\}.*")) {
            return text;
        }

        try {
            StringWriter writer = new StringWriter();
            Template template = new Template("inline", text, cfg);
            template.process(fieldsJson, writer);
            return writer.toString();
        } catch (ParseException e) {
            System.err.println("Erro ao processar template: " + e.getMessage());
            System.err.println("Texto problemático: " + text);
            return text;
        } catch (Exception e) {
            System.err.println("Erro inesperado ao processar template: " + e.getMessage());
            return text;
        }
    }

    private void processParagraphWithFreemarker(
            XWPFParagraph paragraph,
            Configuration cfg,
            Map<String, Object> fieldsJson) throws IOException, TemplateException {
        String text = paragraph.getText();
        String templateName = "template";
        StringTemplateLoader stringLoader = new StringTemplateLoader();
        cfg.setTemplateLoader(stringLoader);
        stringLoader.putTemplate(templateName, text);

        Template template = cfg.getTemplate(templateName);

        Map<String, Object> fieldsCopy = new HashMap<>(fieldsJson);

        StringWriter writer = new StringWriter();
        template.process(fieldsCopy, writer);

        while (paragraph.getRuns().size() > 0) {
            paragraph.removeRun(0);
        }
        paragraph.createRun().setText(writer.toString());
    }
}
