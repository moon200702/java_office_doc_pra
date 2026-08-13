# DOCX Auto-Fill Usage Guide

## Overview
This Java program automatically fills placeholders in .docx (Microsoft Word) files with data from a JSON file.

## Prerequisites
- Java 11 or higher
- Maven 3.6 or higher

## Building the Project

```bash
mvn clean package
```

This creates an executable JAR file in the `target/` directory.

## Usage

### Method 1: Command Line

```bash
java -cp target/docx-autofill-1.0.0.jar com.office.automation.DocxAutoFiller <input.docx> <output.docx> <data.json>
```

**Example:**
```bash
java -cp target/docx-autofill-1.0.0.jar com.office.automation.DocxAutoFiller template.docx result.docx data/sample_data.json
```

### Method 2: Programmatic Usage

```java
import com.office.automation.DocxAutoFiller;
import com.office.automation.DataLoader;
import java.util.Map;

// Load data from JSON
Map<String, String> data = DataLoader.loadFromJson("data.json");

// Or create data map directly
Map<String, String> data = DataLoader.createDataMap(
    "name", "John Doe",
    "email", "john@example.com",
    "date", "2024-01-15"
);

// Fill document
DocxAutoFiller.processDocument("template.docx", "output.docx", data);
```

## Template Creation

### Step 1: Create a Template
1. Open Microsoft Word (or LibreOffice Writer)
2. Create your document structure
3. Add placeholders in the format: `${placeholder_name}`

### Step 2: Example Template Content
```
Dear ${name},

Thank you for your interest in ${company}.

Email: ${email}
Phone: ${phone}
Position: ${position}

Best regards,
HR Department
```

### Step 3: Save as .docx Format
- File → Save As
- Choose "Word Document (.docx)" format

## Data File Format (JSON)

Create a JSON file with key-value pairs matching your placeholders:

```json
{
  "name": "John Smith",
  "email": "john.smith@example.com",
  "company": "ABC Corporation",
  "phone": "+1-555-0123",
  "position": "Senior Developer"
}
```

## Features

### Basic Features
- ✅ Replace placeholders in paragraphs
- ✅ Replace placeholders in table cells
- ✅ Replace placeholders in headers and footers
- ✅ JSON data source support

### Advanced Features (AdvancedDocxFiller)
- ✅ Preserve text formatting (bold, italic, colors, fonts)
- ✅ Batch process multiple documents
- ✅ Handle complex data structures
- ✅ Support for empty/null values

## Examples

### Simple Letter Fill

**template.docx:**
```
Dear ${name},

I hope this message finds you well.

Best regards,
${sender_name}
```

**data.json:**
```json
{
  "name": "Alice Johnson",
  "sender_name": "Bob Wilson"
}
```

**Command:**
```bash
java -cp target/docx-autofill-1.0.0.jar com.office.automation.DocxAutoFiller template.docx output.docx data.json
```

**Result:** 
```
Dear Alice Johnson,

I hope this message finds you well.

Best regards,
Bob Wilson
```

### Invoice Template Fill

**template.docx:**
```
Invoice #${invoice_number}
Date: ${invoice_date}

Bill To:
${customer_name}
${customer_address}

Amount Due: ${total_amount}
```

**data.json:**
```json
{
  "invoice_number": "INV-2024-001",
  "invoice_date": "2024-01-15",
  "customer_name": "John Doe",
  "customer_address": "123 Main St, City, State 12345",
  "total_amount": "$1,500.00"
}
```

### Table Fill

You can also fill table cells - just place placeholders inside table cells:

| Field | Value |
|-------|-------|
| Name | ${name} |
| Email | ${email} |
| Phone | ${phone} |

The program will automatically replace all placeholders in table cells.

## Advanced Usage

### Using AdvancedDocxFiller for Format Preservation

```java
import com.office.automation.AdvancedDocxFiller;
import java.util.Map;

AdvancedDocxFiller filler = new AdvancedDocxFiller();
Map<String, String> data = DataLoader.loadFromJson("data.json");

filler.fillDocumentPreservingFormat("template.docx", "output.docx", data);
```

### Batch Processing

```java
List<String> templates = List.of("template1.docx", "template2.docx", "template3.docx");
filler.batchProcess(templates, "output_dir/", data);
```

## Troubleshooting

### Issue: Placeholders not being replaced
- **Solution**: Ensure placeholder format is exactly `${key_name}` (no spaces)
- Check that keys in JSON match placeholder names exactly (case-sensitive)

### Issue: Document formatting lost
- **Solution**: Use `AdvancedDocxFiller` instead of basic `DocxAutoFiller`

### Issue: Java ClassNotFoundException
- **Solution**: Ensure you're using the correct JAR path in classpath

### Issue: JSON parsing error
- **Solution**: Validate JSON file format - use online JSON validator
- Ensure all strings are properly quoted

## Performance Considerations

- **Small files**: < 1 second per document
- **Large files**: 1-5 seconds per document
- **Batch processing**: Process multiple documents sequentially

## Limitations

- ❌ Macros are not supported
- ❌ Complex VBA scripts will be ignored
- ❌ Some advanced formatting may not be preserved in basic mode
- ❌ Placeholders cannot span multiple runs (formatting breaks)

## Dependencies

- **Apache POI**: For Word document manipulation
- **Gson**: For JSON parsing
- **SLF4J**: For logging

## License

This project is open source and available for use under the MIT License.

## Support

For issues or feature requests, please refer to the project documentation or contact the development team.
