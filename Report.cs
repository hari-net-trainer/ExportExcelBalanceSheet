using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLSample
{
    public class Report
    {
        public void CreateExcelDoc(string fileName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                // Adding style
                WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = GenerateStylesheet();
                stylePart.Stylesheet.Save();

                // Setting up columns
                Columns columns = new Columns(
                    new Column // empty
                    {
                        Min = 1,
                        Max = 1,
                        Width = 8,
                        CustomWidth = true
                    },
                        new Column // Liab
                        {
                            Min = 2,
                            Max = 2,
                            Width = 30,
                            CustomWidth = true
                        },
                        new Column // amt
                        {
                            Min = 3,
                            Max = 3,
                            Width = 15,
                            CustomWidth = true
                        },
                        new Column //amt
                        {
                            Min = 4,
                            Max = 4,
                            Width = 15,
                            CustomWidth = true
                        },
                        new Column // asset
                        {
                            Min = 5,
                            Max = 5,
                            Width = 30,
                            CustomWidth = true
                        },
                        new Column // amt
                        {
                            Min = 6,
                            Max = 6,
                            Width = 15,
                            CustomWidth = true
                        },
                        new Column // amt
                        {
                            Min = 7,
                            Max = 7,
                            Width = 15,
                            CustomWidth = true
                        }
                        );

                worksheetPart.Worksheet.AppendChild(columns);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Balance Sheet" };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();

                var balList = BalanceSheetViewModel.GetFiscalYearBalanaceSheet();

                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // Constructing header
                Row row = new Row();

                row.Append(
                    ConstructMergedCells("", "A1", 7, 0));

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                row = new Row();
                row.Append(
                    ConstructCell("", CellValues.String, 0, "A2"),
                    ConstructCell("Liabilities", CellValues.String, 2, "B2"),
                    ConstructCell("Amount(RM)", CellValues.String, 2, "C2"),
                    ConstructCell("Amount(RM)", CellValues.String, 2, "D2"),
                    ConstructCell("Assets", CellValues.String, 2, "E2"),
                    ConstructCell("Amount(RM)", CellValues.String, 2, "F2"),
                    ConstructCell("Amount(RM)", CellValues.String, 2, "G2"));

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                // Inserting each employee
                //foreach (var employee in employees)
                //{
                //    row = new Row();

                //    row.Append(
                //        ConstructCell(employee.Id.ToString(), CellValues.Number, 1),
                //        ConstructCell(employee.Name, CellValues.String, 1),
                //        ConstructCell(employee.DOB.ToString("yyyy/MM/dd"), CellValues.String, 1),
                //        ConstructCell(employee.Salary.ToString(), CellValues.Number, 1));

                //    sheetData.AppendChild(row);
                //}

                worksheetPart.Worksheet.Save();
            }
        }

        private Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = null;

            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 }

                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 10 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }

                ));

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } })
                    { PatternType = PatternValues.Solid }) // Index 2 - header
                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // body
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true } // header
                );

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }

        private IEnumerable<Cell> ConstructMergedCells(string firstCellValue, string cellRef, int mergeLen, uint styleIndex = 1)
        {
            List<Cell> retCells = new List<Cell>();
            retCells.Add(ConstructCell(firstCellValue, CellValues.String, styleIndex, cellRef));
            var nextCellRef = cellRef;
            for (int i = 1; i < mergeLen; i++)
            {
                nextCellRef = IncrementXLColumn(nextCellRef);
                retCells.Add(ConstructCell(string.Empty, CellValues.String, styleIndex, nextCellRef));
            }

            return retCells;
        }

        private Cell ConstructCell(string value, CellValues dataType, uint styleIndex = 0, string cellReference = "")
        {
            if (string.IsNullOrWhiteSpace(cellReference))
            {
                return new Cell()
                {
                    CellValue = new CellValue(value),
                    DataType = new EnumValue<CellValues>(dataType),
                    StyleIndex = styleIndex
                };
            }
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                StyleIndex = styleIndex,
                CellReference = cellReference
            };
        }

        private string IncrementXLColumn(string Address, int lentght = 1)
        {
            var parts = System.Text.RegularExpressions.Regex.Matches(Address, @"([A-Z]+)|(\d+)");
            if (parts.Count != 2) return null;
            var tempLen = lentght == 1 ? lentght : lentght - 1;
            return incCol(parts[0].Value, tempLen) + parts[1].Value;
        }

        private string incCol(string col, int lentght = 1)
        {
            if (col == "")
                return "A";
            string fPart = col.Substring(0, col.Length - 1);
            char lChar = col[col.Length - 1];
            if (lChar == 'Z')
            {
                var nextStr = string.IsNullOrWhiteSpace(fPart) ? "A" : incCol(fPart);
                return incCol(nextStr + "A", lentght - 1);
            }
            return fPart + (char)(lChar + lentght);
        }
    }
}
