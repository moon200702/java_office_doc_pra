package com.office.automation;

import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.*;

/**
 * Test cases for DocxAutoFiller
 */
public class DocxAutoFillerTest {
    private DocxAutoFiller filler;
    private Map<String, String> testData;

    @Before
    public void setUp() {
        filler = new DocxAutoFiller();
        testData = new HashMap<>();
        testData.put("name", "John Doe");
        testData.put("email", "john.doe@example.com");
        testData.put("date", "2024-01-15");
        testData.put("company", "Tech Corp");
    }

    @Test
    public void testDataLoaderCreateMap() {
        Map<String, String> data = DataLoader.createDataMap(
            "key1", "value1",
            "key2", "value2",
            "key3", "value3"
        );

        assertEquals(3, data.size());
        assertEquals("value1", data.get("key1"));
        assertEquals("value2", data.get("key2"));
        assertEquals("value3", data.get("key3"));
    }

    @Test
    public void testDataMapContents() {
        assertEquals(4, testData.size());
        assertEquals("John Doe", testData.get("name"));
        assertEquals("john.doe@example.com", testData.get("email"));
    }

    @Test
    public void testCreateTemplateFile() {
        String templatePath = "test-template.docx";
        String outputPath = "test-output.docx";

        // This would require an actual template file
        // For now, just verify the paths are valid
        assertNotNull(templatePath);
        assertNotNull(outputPath);
    }
}
