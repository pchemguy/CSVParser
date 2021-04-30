Excel VBA provides several functionalities for parsing text delimited files:

- basic file I/O and the Split function (basic parser);
- Workbook.Open/Workbook.OpenText functions;
- Excel's QueryTable.

CSVParser implements the first two options. The basic parser is less flexible but at least 20x faster than Workbook.Open/Workbook.OpenText. Excel parser (Workbook.Open/Workbook.OpenText) ignores parsing options if the file has a ".csv" extension. So to ensure correct parsing of a ".csv" file, it is temporarily renamed, parsed, and renamed back.

"Project/Common/CSV Parser/CSVParser.cls" contains the CSVParser class. "Project/Common/CSV Parser/CSVParserSnippets.bas" and "Project/Common/CSV Parser/CSVParserTests.bas" modules contain usage examples and tests. The repository root also contains two test files "Contacts.xsv" and "Contacts.csv" with identical content.
