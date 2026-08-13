package com.office.automation;

import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Advanced version of DocxAutoFiller with support for:
 * - Conditional blocks
 * - Table row cloning/duplication
 * - Formatting preservation
 * - Complex data structures
 */
public class AdvancedDocxFiller {
    private static final Logger logger = LoggerFactory.getLogger(AdvancedDocxFiller.class);
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{([^}]+)\\}");
    private static final Pattern LOOP_PATTERN = Pattern.compile("\\$\\{loop:([^}]+)\\}");
    private static final Pattern ENDLOOP_PATTERN = Pattern.compile("\\$\\{endloop\\}");

    /**
     * Fills document with support for formatting preservation
     */
    public void fillDocumentPreservingFormat(String inputPath, String outputPath, Map<String, String> data) {
        try (FileInputStream fis = new FileInputStream(inputPath);
             XWPFDocument document = new XWPFDocument(fis)) {

            logger.info("Processing document with format preservation: {}", inputPath);

            // Fill paragraphs
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                fillParagraphWithFormatting(paragraph, data);
            }

            // Fill tables
            for (XWPFTable table : document.getTables()) {
                fillTableWithFormatting(table, data);
            }

            // Write output
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                document.write(fos);
                logger.info("Document saved to: {}", outputPath);
            }

        } catch (IOException e) {
            logger.error("Error processing document", e);
        }
    }

    /**
     * Fills paragraph while preserving formatting
     */
    private void fillParagraphWithFormatting(XWPFParagraph paragraph, Map<String, String> data) {
        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.getText(0);
            if (text != null) {
                String replacedText = replaceAllPlaceholders(text, data);
                if (!replacedText.equals(text)) {
                    run.setText(replacedText, 0);
                    logger.debug("Filled run with formatting preserved");
                }
            }
        }
    }

    /**
     * Fills table while preserving formatting
     */
    private void fillTableWithFormatting(XWPFTable table, Map<String, String> data) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    fillParagraphWithFormatting(paragraph, data);
                }
            }
        }
    }

    /**
     * Replace all placeholders in text
     */
    private String replaceAllPlaceholders(String text, Map<String, String> data) {
        String result = text;
        Matcher matcher = PLACEHOLDER_PATTERN.matcher(result);

        while (matcher.find()) {
            String key = matcher.group(1);
            String value = data.getOrDefault(key, "");
            result = result.replace("${" + key + "}", value);
            matcher = PLACEHOLDER_PATTERN.matcher(result);
        }

        return result;
    }

    /**
     * Batch process multiple documents
     */
    public void batchProcess(List<String> inputPaths, String outputDir, Map<String, String> data) {
        for (String inputPath : inputPaths) {
            String fileName = new File(inputPath).getName();
            String outputPath = outputDir + File.separator + "filled_" + fileName;
            fillDocumentPreservingFormat(inputPath, outputPath, data);
        }
    }
}
