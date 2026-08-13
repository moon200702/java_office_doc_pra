# DOCX Auto-Fill (Simple Guide)

A small Java tool that fills placeholders in Microsoft Word (.docx) files using data from a JSON file. Use it to create letters, invoices, certificates, and other documents automatically.

## Quick overview

- The program finds placeholders in a Word document that look like `${key}`.
- It replaces each placeholder with the matching value from a JSON file.
- Works for text in paragraphs, tables, headers, and footers.
- There is a simple mode and an advanced mode that keeps text formatting.

## Requirements

- Java 11 or newer
- Maven 3.6 or newer
- Windows, macOS, or Linux

## Build

Clone the repo and build with Maven:

```bash
git clone https://github.com/moon200702/java_office_doc_pra.git
cd java_office_doc_pra
mvn clean package -DskipTests
```

This creates a standalone JAR in `target/`:
- `docx-autofill-jar-with-dependencies.jar` — recommended, contains all libraries.

## Command-line usage

Run the program with:

```bash
java -jar target/docx-autofill-jar-with-dependencies.jar <input.docx> <output.docx> <data.json>
```

Example:

```bash
java -jar target/docx-autofill-jar-with-dependencies.jar template.docx result.docx data/sample_data.json
```

Arguments:
- `input.docx` — your template file with placeholders like `${name}`
- `output.docx` — file to write with filled values
- `data.json` — JSON file containing keys and values

## JSON format

Use a simple object with string values. Example `data/sample_data.json`:

```json
{
  "name": "John Smith",
  "company": "ABC Corporation",
  "date": "2024-01-15",
  "email": "john@example.com"
}
```

Placeholders are case-sensitive and must match the JSON keys exactly.

## Simple Java usage

If you want to call the filler from Java code:

```java
import com.office.automation.DocxAutoFiller;
import com.office.automation.DataLoader;
import java.util.Map;

Map<String, String> data = DataLoader.loadFromJson("data/sample_data.json");
DocxAutoFiller.processDocument("template.docx", "output.docx", data);
```

## Advanced usage (format-preserving)

Use AdvancedDocxFiller to preserve bold/italic/font styles and for batch processing:

```java
import com.office.automation.AdvancedDocxFiller;
import com.office.automation.DataLoader;
import java.util.List;
import java.util.Map;

AdvancedDocxFiller filler = new AdvancedDocxFiller();
Map<String, String> data = DataLoader.loadFromJson("data/sample_data.json");
filler.fillDocumentPreservingFormat("template.docx", "output.docx", data);

// Batch example
List<String> templates = List.of("t1.docx", "t2.docx");
filler.batchProcess(templates, "output_dir/", data);
```

## Common problems & fixes

- Placeholders not replaced:
  - Make sure placeholder is exactly `${key}`.
  - JSON key must match placeholder name (case sensitive).
  - If the placeholder has different formatting inside Word, use the advanced filler or clear formatting in Word.

- JSON errors:
  - Validate your JSON (no trailing commas, proper quotes).

- Missing classes (like SLF4J):
  - Use the fat JAR `docx-autofill-jar-with-dependencies.jar`.

- Large documents:
  - Increase Java heap: `java -Xmx2g -jar ...`

## Notes & tips

- Keep placeholders simple and consistent (e.g., `${first_name}`).
- Prefer the advanced filler when templates use varied formatting.
- Test with a copy of your template first.

## Libraries used

- Apache POI (docx handling)
- Gson (JSON parsing)
- SLF4J (logging)

---

If you want, I can:
- shorten this further into a one-page quick-start,
- add screenshots or sample files,
- or produce a printable quick-reference sheet.
```
