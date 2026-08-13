# DOCX Auto-Fill - Java Office Document Automation

A powerful Java application for automatically filling placeholders in Microsoft Word (.docx) documents with data from JSON files. Perfect for generating personalized documents, certificates, invoices, and more.

## 📋 Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation & Build](#installation--build)
- [Usage](#usage)
- [Library Details](#library-details)
- [Project Structure](#project-structure)
- [Examples](#examples)
- [Advanced Features](#advanced-features)
- [Troubleshooting](#troubleshooting)

## ✨ Features

- ✅ **Placeholder Replacement**: Replace `${placeholder_name}` with actual data
- ✅ **Multi-Element Support**: Works in paragraphs, tables, headers, and footers
- ✅ **JSON Data Source**: Load data from JSON files
- ✅ **Format Preservation**: Advanced mode preserves text formatting (bold, italic, colors, fonts)
- ✅ **Batch Processing**: Process multiple documents in one run
- ✅ **Standalone JAR**: Single JAR file with all dependencies included
- ✅ **Comprehensive Logging**: SLF4J logging for debugging
- ✅ **Exception Handling**: Robust error handling and reporting

## 🔧 Requirements

| Requirement | Version |
|-------------|---------|
| Java | 11 or higher |
| Maven | 3.6 or higher |
| Operating System | Windows, macOS, Linux |

## 📦 Installation & Build

### Clone the Repository

```bash
git clone https://github.com/moon200702/java_office_doc_pra.git
cd java_office_doc_pra
```

### Build with Maven

```bash
mvn clean package -DskipTests
```

This creates two JAR files in `target/`:
- `docx-autofill-1.0.0.jar` (11 KB) - Requires dependencies on classpath
- `docx-autofill-jar-with-dependencies.jar` (18 MB) - **Recommended** - Standalone

## 🚀 Usage

### Command Line (Recommended)

```bash
java -jar target/docx-autofill-jar-with-dependencies.jar <input.docx> <output.docx> <data.json>
```

**Parameters:**
- `input.docx` - Path to template Word document
- `output.docx` - Path to output filled document
- `data.json` - Path to JSON file with placeholder values

**Example:**

```bash
java -jar target/docx-autofill-jar-with-dependencies.jar template.docx result.docx data/sample_data.json
```

### Programmatic Usage (Java Code)

#### Basic Usage

```java
import com.office.automation.DocxAutoFiller;
import com.office.automation.DataLoader;
import java.util.Map;

// Load data from JSON file
Map<String, String> data = DataLoader.loadFromJson("data/sample_data.json");

// Fill document
DocxAutoFiller.processDocument("template.docx", "output.docx", data);
```

#### Direct Map Creation

```java
import com.office.automation.DocxAutoFiller;
import com.office.automation.DataLoader;
import java.util.Map;

// Create data map directly
Map<String, String> data = DataLoader.createDataMap(
    "name", "John Smith",
    "email", "john@example.com",
    "date", "2024-01-15",
    "company", "ABC Corporation"
);

// Fill document
DocxAutoFiller.processDocument("template.docx", "output.docx", data);
```

#### Advanced Usage with Format Preservation

```java
import com.office.automation.AdvancedDocxFiller;
import com.office.automation.DataLoader;
import java.util.List;
import java.util.Map;

AdvancedDocxFiller filler = new AdvancedDocxFiller();
Map<String, String> data = DataLoader.loadFromJson("data/sample_data.json");

// Fill with format preservation
filler.fillDocumentPreservingFormat("template.docx", "output.docx", data);

// Batch processing
List<String> templates = List.of("template1.docx", "template2.docx");
filler.batchProcess(templates, "output_dir/", data);
```

## 📚 Library Details

### Dependencies

All dependencies are automatically included in the fat JAR. Here are the libraries used:

#### Core Libraries

| Library | Version | Purpose |
|---------|---------|---------|
| **Apache POI OOXML** | 5.2.3 | Reads and writes Microsoft Office documents (.docx) |
| **Gson** | 2.10.1 | JSON parsing and serialization |
| **SLF4J API** | 2.0.7 | Logging abstraction layer |
| **SLF4J Simple** | 2.0.7 | Simple logging implementation |

#### Transitive Dependencies (included via POI)

- `poi-ooxml-lite` - Optimized POI libraries
- `xmlbeans` - XML processing for Office documents
- `commons-compress` - Archive compression utilities
- `commons-io` - I/O utilities
- `commons-collections` - Collection utilities

### Maven Build Plugins

| Plugin | Version | Purpose |
|--------|---------|---------|
| maven-compiler-plugin | 3.11.0 | Java compilation (source: Java 11) |
| maven-jar-plugin | 3.3.0 | JAR packaging |
| maven-assembly-plugin | 3.6.0 | Fat JAR creation with dependencies |

### Maven Configuration

#### Dependencies in pom.xml

```xml
<!-- Apache POI for Word document manipulation -->
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.2.3</version>
</dependency>

<!-- Logging -->
<dependency>
    <groupId>org.slf4j</groupId>
    <artifactId>slf4j-api</artifactId>
    <version>2.0.7</version>
</dependency>
<dependency>
    <groupId>org.slf4j</groupId>
    <artifactId>slf4j-simple</artifactId>
    <version>2.0.7</version>
</dependency>

<!-- JSON processing -->
<dependency>
    <groupId>com.google.code.gson</groupId>
    <artifactId>gson</artifactId>
    <version>2.10.1</version>
</dependency>

<!-- Testing -->
<dependency>
    <groupId>junit</groupId>
    <artifactId>junit</artifactId>
    <version>4.13.2</version>
    <scope>test</scope>
</dependency>
```

## 📁 Project Structure

```
java_office_doc_pra/
├── src/
│   ├── main/
│   │   └── java/com/office/automation/
│   │       ├── DocxAutoFiller.java           # Main class for basic filling
│   │       ├── AdvancedDocxFiller.java      # Advanced class with format preservation
│   │       └── DataLoader.java              # Utility for loading data
│   └── test/
│       └── java/com/office/automation/
│           └── DocxAutoFillerTest.java      # Unit tests
├── data/
│   └── sample_data.json                     # Sample data file
├── pom.xml                                   # Maven configuration
├── README.md                                 # This file
├── USAGE.md                                  # Detailed usage guide
└── .gitignore                                # Git ignore rules
```

## 📝 Examples

### Example 1: Personal Letter

**Template (template.docx):**
```
Dear ${name},

I hope this message finds you well. We are pleased to confirm your position as 
${position} at ${company}.

Your email: ${email}
Your phone: ${phone}

Best regards,
HR Department
```

**Data (data/sample_data.json):**
```json
{
  "name": "John Smith",
  "position": "Senior Developer",
  "company": "Tech Corporation",
  "email": "john.smith@techcorp.com",
  "phone": "+1-555-0123"
}
```

**Command:**
```bash
java -jar target/docx-autofill-jar-with-dependencies.jar template.docx letter.docx data/sample_data.json
```

**Result:** All placeholders replaced with actual values ✅

### Example 2: Invoice Template

**Template (invoice_template.docx):**
```
INVOICE

Invoice #: ${invoice_number}
Date: ${invoice_date}
Due Date: ${due_date}

Bill To:
${customer_name}
${customer_address}
${customer_city}, ${customer_state} ${customer_zip}

Items:
[Table with ${item_1}, ${item_2}, etc.]

Subtotal: ${subtotal}
Tax: ${tax}
Total: ${total}
```

**Data (data/invoice_data.json):**
```json
{
  "invoice_number": "INV-2024-001",
  "invoice_date": "2024-01-15",
  "due_date": "2024-02-15",
  "customer_name": "ABC Company",
  "customer_address": "123 Business Ave",
  "customer_city": "New York",
  "customer_state": "NY",
  "customer_zip": "10001",
  "item_1": "Consulting Services",
  "item_2": "Software License",
  "subtotal": "$5,000.00",
  "tax": "$500.00",
  "total": "$5,500.00"
}
```

### Example 3: Certificate

**Template (certificate_template.docx):**
```
Certificate of Completion

This certifies that ${name} has successfully completed 
${course_name} on ${completion_date}.

Instructor: ${instructor_name}
```

**Data (data/certificate_data.json):**
```json
{
  "name": "Alice Johnson",
  "course_name": "Advanced Java Programming",
  "completion_date": "January 15, 2024",
  "instructor_name": "Dr. Bob Wilson"
}
```

## 🔬 Advanced Features

### Batch Processing Multiple Documents

```java
AdvancedDocxFiller filler = new AdvancedDocxFiller();
List<String> templates = Arrays.asList(
    "template1.docx",
    "template2.docx", 
    "template3.docx"
);
Map<String, String> data = DataLoader.loadFromJson("data.json");

filler.batchProcess(templates, "output/", data);
```

### Format Preservation

The `AdvancedDocxFiller` class preserves formatting:
- Text styles (bold, italic, underline)
- Font families and sizes
- Text colors
- Paragraph formatting

```java
AdvancedDocxFiller filler = new AdvancedDocxFiller();
filler.fillDocumentPreservingFormat("template.docx", "output.docx", data);
```

## 🐛 Troubleshooting

### Issue: ClassNotFoundException - Missing Dependencies

**Error:**
```
java.lang.ClassNotFoundException: org/slf4j/LoggerFactory
```

**Solution:** Use the fat JAR with dependencies:
```bash
java -jar target/docx-autofill-jar-with-dependencies.jar ...
```

### Issue: Placeholders Not Being Replaced

**Possible Causes:**
1. Placeholder format is incorrect (must be exactly `${key_name}`)
2. JSON key doesn't match placeholder name (case-sensitive)
3. Placeholder spans multiple runs in Word (formatting breaks it)

**Solution:** 
- Verify placeholder format in template
- Check JSON keys match exactly
- In Word, select placeholder text and use "Clear Formatting"

### Issue: JSON Parsing Error

**Error:**
```
com.google.gson.JsonSyntaxException: ...
```

**Solution:** Validate JSON file using online JSON validator:
- Ensure all strings are quoted
- Check for trailing commas
- Verify proper nesting

### Issue: Output File Is Corrupted

**Possible Cause:** Placeholder spans multiple runs/lines in Word

**Solution:** Use `AdvancedDocxFiller` with format preservation:
```java
AdvancedDocxFiller filler = new AdvancedDocxFiller();
filler.fillDocumentPreservingFormat("template.docx", "output.docx", data);
```

## 📊 Performance Benchmarks

| Document Size | Processing Time | Notes |
|---------------|-----------------|-------|
| Small (< 1 MB) | < 1 second | Simple letters, templates |
| Medium (1-5 MB) | 1-3 seconds | Documents with tables |
| Large (5+ MB) | 3-10 seconds | Complex documents, many images |

## 🔒 Limitations

- ❌ Macros not supported (will be ignored)
- ❌ Complex VBA scripts not executed
- ❌ Some advanced formatting may not be preserved in basic mode
- ❌ Placeholders cannot span multiple formatting runs
- ⚠️ Very large documents (100+ MB) may require increased heap memory

## 💡 Tips & Best Practices

1. **Template Creation**: Use a clear naming convention for placeholders (e.g., `${customer_first_name}` instead of `${cfn}`)

2. **Data Preparation**: Always validate JSON data before running bulk operations

3. **Formatting**: Keep template formatting simple to avoid placeholder spanning issues

4. **Error Handling**: Check logs for detailed information about processing

5. **Memory**: For large files, increase Java heap:
   ```bash
   java -Xmx2g -jar target/docx-autofill-jar-with-dependencies.jar ...
   ```

## 📄 License

Open source - available for personal and commercial use

## 👨‍💻 Contributing

Contributions welcome! Please feel free to submit pull requests or report issues.

## 📧 Support

For issues or questions, please refer to the USAGE.md file for more detailed documentation.