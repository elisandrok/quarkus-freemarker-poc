package com.viacerta.service;

import jakarta.enterprise.context.ApplicationScoped;

import java.util.List;
import java.util.Map;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;

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

        addPageNumbering(document);
    }

    private void processHeaders(XWPFDocument document, Map<String, String> fields) {
        for (XWPFHeader header : document.getHeaderList()) {
            for (XWPFParagraph paragraph : header.getParagraphs()) {
                processHeaderParagraph(paragraph, fields);
            }
        }
    }

    private void processHeaderParagraph(XWPFParagraph paragraph, Map<String, String> fields) {
        List<XWPFRun> runs = paragraph.getRuns();
        for (int i = 0; i < runs.size(); i++) {
            XWPFRun run = runs.get(i);
            if (run.getEmbeddedPictures().size() > 0) {
                // Este run contém uma imagem, não o modificamos
                continue;
            }
            if (run.getText(0) != null) {
                String text = run.getText(0);
                String newText = replaceFields(text, fields);
                if (!text.equals(newText)) {
                    // Criar um novo run com o texto substituído
                    XWPFRun newRun = paragraph.insertNewRun(i);
                    newRun.setText(newText);
                    copyRunProperties(run, newRun);
                    
                    // Remover o run original
                    paragraph.removeRun(i + 1);
                }
            }
        }
    }

    private void processParagraph(XWPFParagraph paragraph, Map<String, String> fields, Map<String, Boolean> rules) {
        String paragraphText = paragraph.getText();
        if (rules != null && shouldRemoveParagraph(paragraph, rules)) {
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

    public void addPageNumbering(XWPFDocument document) {
        // Criar ou obter o rodapé para todas as páginas
        XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);
        
        // Criar um parágrafo no rodapé
        XWPFParagraph paragraph = footer.getParagraphArray(0);
        if (paragraph == null) {
            paragraph = footer.createParagraph();
        }
        paragraph.setAlignment(ParagraphAlignment.RIGHT);

        // Adicionar o texto "Página "
        XWPFRun run = paragraph.createRun();
        run.setText("Página ");

        // Adicionar o campo de número da página atual
        insertField(paragraph, "PAGE");

        run = paragraph.createRun();
        run.setText(" de ");

        // Adicionar o campo de número total de páginas
        insertField(paragraph, "NUMPAGES");
    }

    private void insertField(XWPFParagraph paragraph, String fieldName) {
        CTSimpleField field = paragraph.getCTP().addNewFldSimple();
        field.setInstr(" " + fieldName + " ");
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
        targetRun.setTextPosition(sourceRun.getTextPosition());
        // Adicionar outras propriedades caso necessário
    }
}
