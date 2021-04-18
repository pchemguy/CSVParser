## CSVParser

Implements two CSV parsing approaches for delimited files: pure VBA (very basic parser) based on the "Split" function, and Excel assisted parsing using either `Workbook.Open` or `Workbook.OpenText`. While the Excel parser is more flexible, it is 1-2 orders of magnitude slower than the basic parser. Excel parser ignores parsing options if the file has a ".csv" extension. So to ensure correct parsing, if the file provided has a ".csv" extension, it is temporarily renamed, parsed, and renamed back.

Main module - "Project/Common/CSV Parser/CSVParser.cls" - `CSVParser` class. Usage examples - "Project/Common/CSV Parser/CSVParserSnippets.bas" and "Project/Common/CSV Parser/CSVParserTests.bas".
