package com.office.automation;

import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Main class for auto-filling .docx files with data.
 * Supports placeholder replacement in paragraphs, tables, and headers/footers.
 */
public class DocxAutoFiller {
    private static final Logger logger = LoggerFactory.getLogger(DocxAutoFiller.class);
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{([^}]+)\\}");

    /**
     * Auto-fills a .docx document with data from a map
     *
     * @param inputPath  Path to input .docx file
     * @param outputPath Path to output filled .docx file
     * @param data       Map containing placeholder-value pairs
     */
    public void fillDocument(String inputPath, String outputPath, Map<String, String> data) {
        try (FileInputStream fis = new FileInputStream(inputPath);
             XWPFDocument document = new XWPFDocument(fis)) {

            logger.info("Processing document: {}", inputPath);

            // Fill paragraphs
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                fillParagraph(paragraph, data);
            }

            // Fill tables
            for (XWPFTable table : document.getTables()) {
                fillTable(table, data);
            }

            // Fill headers
            for (XWPFHeader header : document.getHeaderList()) {
                for (XWPFParagraph paragraph : header.getParagraphs()) {
                    fillParagraph(paragraph, data);
                }
            }

            // Fill footers
            for (XWPFFooter footer : document.getFooterList()) {
                for (XWPFParagraph paragraph : footer.getParagraphs()) {
                    fillParagraph(paragraph, data);
                }
            }

            // Write output
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                document.write(fos);
                logger.info("Document saved to: {}", outputPath);
            }

        } catch (FileNotFoundException e) {
            logger.error("Input file not found: {}", inputPath, e);
        } catch (IOException e) {
            logger.error("Error processing document", e);
        }
    }

    /**
     * Fills placeholders in a paragraph
     */
    private void fillParagraph(XWPFParagraph paragraph, Map<String, String> data) {
        String text = paragraph.getText();
        Matcher matcher = PLACEHOLDER_PATTERN.matcher(text);

        if (matcher.find()) {
            // Clear existing runs
            while (paragraph.getRuns().size() > 0) {
                paragraph.removeRun(0);
            }

            // Process text and reconstruct with data
            String replacedText = text;
            for (Map.Entry<String, String> entry : data.entrySet()) {
                String placeholder = "${" + entry.getKey() + "}";
                replacedText = replacedText.replace(placeholder, entry.getValue() != null ? entry.getValue() : "");
            }

            // Add replaced text as new run
            if (!replacedText.isEmpty()) {
                XWPFRun run = paragraph.createRun();
                run.setText(replacedText);
            }

            logger.debug("Filled paragraph: {}", replacedText);
        }
    }

    /**
     * Fills placeholders in table cells
     */
    private void fillTable(XWPFTable table, Map<String, String> data) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    fillParagraph(paragraph, data);
                }
            }
        }
    }

    /**
     * Standalone method to process a single document
     */
    public static void processDocument(String inputPath, String outputPath, Map<String, String> data) {
        DocxAutoFiller filler = new DocxAutoFiller();
        filler.fillDocument(inputPath, outputPath, data);
    }

    public static void main(String[] args) {
        if (args.length < 3) {
            System.out.println("Usage: java DocxAutoFiller <input.docx> <output.docx> <data_json_file>");
            System.out.println("Example: java DocxAutoFiller template.docx output.docx data.json");
            System.exit(1);
        }

        String inputPath = args[0];
        String outputPath = args[1];
        String dataFile = args[2];

        try {
            Map<String, String> data = DataLoader.loadFromJson(dataFile);
            processDocument(inputPath, outputPath, data);
            logger.info("Auto-fill completed successfully!");
        } catch (IOException e) {
            logger.error("Failed to load data file: {}", dataFile, e);
            System.exit(1);
        }
    }
}
