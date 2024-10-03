package com.viacerta.service;

import jakarta.enterprise.context.ApplicationScoped;

import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

@ApplicationScoped
public class FreemarkerProcessorService {
    public void processTemplate(XWPFDocument document, Map<String, String> fields, Map<String, Boolean> rules) {
        // Processar cabeçalhos
        processHeaders(document, fields);
        // Processar parágrafos
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            processParagraph(paragraph, fields, rules);
        }

        // Processar tabelas
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        processParagraph(paragraph, fields, rules);
                    }
                }
            }
        }
    }

    private void processHeaders(XWPFDocument document, Map<String, String> fields) {
        for (XWPFHeader header : document.getHeaderList()) {
            for (XWPFParagraph paragraph : header.getParagraphs()) {
                processHeaderParagraph(paragraph, fields);
            }
        }
    }

    private void processHeaderParagraph(XWPFParagraph paragraph, Map<String, String> fields) {
        String fullText = paragraph.getText();
        String newText = replaceFields(fullText, fields);
        
        if (!fullText.equals(newText)) {
            // Limpar o parágrafo existente
            while (paragraph.getRuns().size() > 0) {
                paragraph.removeRun(0);
            }
            
            // Adicionar o novo texto preservando a formatação original
            XWPFRun run = paragraph.createRun();
            run.setText(newText);
            copyRunProperties(paragraph.getRuns().get(0), run);
        }
    }

    private void processParagraph(XWPFParagraph paragraph, Map<String, String> fields, Map<String, Boolean> rules) {
        String paragraphText = paragraph.getText();
        if (shouldRemoveParagraph(paragraph, rules)) {
            while (paragraph.getRuns().size() > 0) {
                paragraph.removeRun(0);
            }
            return;
        }

        StringBuilder paragraphTextBuilder = new StringBuilder();
        List<XWPFRun> runs = paragraph.getRuns();
        
        // Coletar todo o texto do parágrafo
        for (XWPFRun run : runs) {
            String runText = run.getText(0);
            if (runText != null) {
                paragraphTextBuilder.append(runText);
            }
        }

        // Substituir os placeholders no texto completo do parágrafo
        String newText = replaceFields(paragraphText.toString(), fields);

        // Se o texto foi modificado, atualizar os runs
        if (!paragraphText.toString().equals(newText)) {
            // Limpar todos os runs existentes
            while (paragraph.getRuns().size() > 0) {
                paragraph.removeRun(0);
            }

            // Criar um novo run com o texto atualizado
            XWPFRun newRun = paragraph.createRun();
            newRun.setText(newText);

            // Copiar a formatação do primeiro run original para o novo run
            if (!runs.isEmpty()) {
                copyRunProperties(runs.get(0), newRun);
            }
        }
    }

    private boolean shouldRemoveParagraph(XWPFParagraph paragraph, Map<String, Boolean> rules) {
        if (rules == null) {
            return false;
        }
        String text = paragraph.getText();

        // Verificar condições do tipo ${if(...)}
        if (text.contains("${if(")) {
            int start = text.indexOf("${if(") + 5;
            int end = text.indexOf(")}");
            if (end > start) {
                String condition = text.substring(start, end);
                if (condition.contains("<>")) {
                    String[] parts = condition.split("<>");
                    if (parts.length == 2) {
                        String fieldName = parts[0].trim();
                        //String value = parts[1].trim().replace("\"", "");
                        Boolean ruleValue = rules.get(fieldName);
                        return ruleValue == null || !ruleValue;
                    }
                }
            }
        }
        
        // Verificar regras normais
        for (Map.Entry<String, Boolean> rule : rules.entrySet()) {
            if (text.contains(rule.getKey()) && !rule.getValue()) {
                return true;
            }
        }
        return false;
    }

    private String replaceFields(String text, Map<String, String> fields) {
        for (Map.Entry<String, String> entry : fields.entrySet()) {
            String placeholder = "{{" + entry.getKey() + "}}";
            String value = entry.getValue() != null ? entry.getValue() : "";
            text = text.replace(placeholder, value);
        }
        return text;
    }

    private void copyRunProperties(XWPFRun sourceRun, XWPFRun targetRun) {
        targetRun.setBold(sourceRun.isBold());
        targetRun.setItalic(sourceRun.isItalic());
        targetRun.setUnderline(sourceRun.getUnderline());
        targetRun.setColor(sourceRun.getColor());
        targetRun.setFontFamily(sourceRun.getFontFamily());
        Double fontSize = sourceRun.getFontSizeAsDouble();
        if (fontSize != null) {
            targetRun.setFontSize(fontSize);
        }
        // Adicionar outras propriedades caso necessário
    }
}
