
## CSVParser

Implements two CSV parsing approaches for delimited files: pure VBA (very basic parser) based on "Split" function, and Excel assisted parsing using either `Workboon.Open` or `Workboon.OpenText`. Excel parser ignores parsing options if the file has ".csv" extension. So to ensure correct parsing, if the file provided has a ".csv" extension, it is temporarily renamed, parsed, and renamed back.

Main module - "Project/Common/CSV Parser/CSVParser.cls" - `CSVParser` class. Usage examples - "Project/Common/CSV Parser/CSVParserSnippets.bas" and "Project/Common/CSV Parser/CSVParserTests.bas".