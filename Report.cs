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
        private Dictionary<UInt32, Row> RowDicSheetData = new Dictionary<uint, Row>();
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

                //create a MergeCells class to hold each MergeCell
                MergeCells mergeCells = new MergeCells();
                UInt32 rowIndex = 1;
                var refCell = string.Empty;
                string mergeRef = string.Empty;

                // Constructing rows
                Row row = new Row();
                //empty row
                row.RowIndex = rowIndex;
                row.Append(
                    ConstructMergedCells(string.Empty, string.Format("{0}{1}", "A", rowIndex), 7, 0));
                sheetData.AppendChild(row);
                rowIndex++;

                //condo name - merge and center, all borders
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)));
                refCell = string.Format("{0}{1}", "B", rowIndex);
                row.Append(
                    ConstructMergedCells(balList.CondoName, refCell, 6, 6));
                sheetData.AppendChild(row);
                rowIndex++;
                mergeRef = IncrementXLColumn(refCell, 6);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });

                //condo Add - merge and center, all borders
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)));
                refCell = string.Format("{0}{1}", "B", rowIndex);
                row.Append(
                    ConstructMergedCells(balList.CondoAdd, refCell, 6, 6));
                sheetData.AppendChild(row);
                rowIndex++;
                mergeRef = IncrementXLColumn(refCell, 6);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });

                //Fiscal year period - merge and center, all borders
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)));
                refCell = string.Format("{0}{1}", "B", rowIndex);
                row.Append(
                    ConstructMergedCells("Balance Sheet at \"" + balList.FYPeriod + "\"", refCell, 6, 6));
                sheetData.AppendChild(row);
                rowIndex++;
                mergeRef = IncrementXLColumn(refCell, 6);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });

                //liable&asset col
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(
                    ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)),
                    ConstructCell("Liabilities", CellValues.String, 6, string.Format("{0}{1}", "B", rowIndex)),
                    ConstructCell("Amount(RM)", CellValues.String, 6, string.Format("{0}{1}", "C", rowIndex)),
                    ConstructCell("Amount(RM)", CellValues.String, 6, string.Format("{0}{1}", "D", rowIndex)),
                    ConstructCell("Assets", CellValues.String, 6, string.Format("{0}{1}", "E", rowIndex)),
                    ConstructCell("Amount(RM)", CellValues.String, 6, string.Format("{0}{1}", "F", rowIndex)),
                    ConstructCell("Amount(RM)", CellValues.String, 6, string.Format("{0}{1}", "G", rowIndex)));
                sheetData.AppendChild(row);
                rowIndex++;

                //empty row
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(
                    ConstructMergedCells(string.Empty, string.Format("{0}{1}", "A", rowIndex), 7, 8));
                sheetData.AppendChild(row);
                rowIndex++;

                //data liab in B,C,D and asset in E,F,G
                int liabCount = 0, assetCount = 0;
                UInt32 assetRowIndx = rowIndex;
                RowDicSheetData = new Dictionary<UInt32, Row>();
                //Liabilities
                foreach (var liabItem in balList.LiabilityItem.Liabilities)
                {
                    //catagory Name
                    row = new Row();
                    row.RowIndex = rowIndex;
                    row.Append(
                    ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)),
                    ConstructCell(liabItem.CatagoryItem, CellValues.String, 9, string.Format("{0}{1}", "B", rowIndex)),
                    ConstructCell(string.Empty, CellValues.String, 8, string.Format("{0}{1}", "C", rowIndex)),
                    ConstructCell(string.Empty, CellValues.String, 8, string.Format("{0}{1}", "D", rowIndex)));
                    RowDicSheetData.Add(rowIndex, row);
                    liabCount++;
                    rowIndex++;
                    foreach (var item in liabItem.Items)
                    {
                        //Sub Catagory Names with amount
                        row = new Row();
                        row.RowIndex = rowIndex;
                        row.Append(
                        ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)),
                        ConstructCell(item.ItemName, CellValues.String, 8, string.Format("{0}{1}", "B", rowIndex)),
                        ConstructCell(item.Value, CellValues.String, 7, string.Format("{0}{1}", "C", rowIndex)),
                        ConstructCell(string.Empty, CellValues.String, 8, string.Format("{0}{1}", "D", rowIndex)));
                        RowDicSheetData.Add(rowIndex, row);
                        liabCount++;
                        rowIndex++;
                    }
                    //total amount per catagory
                    row = new Row();
                    row.RowIndex = rowIndex;
                    row.Append(
                    ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)),
                    ConstructCell(string.Empty, CellValues.String, 8, string.Format("{0}{1}", "B", rowIndex)),
                    ConstructCell(string.Empty, CellValues.String, 8, string.Format("{0}{1}", "C", rowIndex)),
                    ConstructCell(liabItem.TotalValue, CellValues.String, 7, string.Format("{0}{1}", "D", rowIndex)));
                    RowDicSheetData.Add(rowIndex, row);
                    liabCount++;
                    rowIndex++;
                }

                //Assets asset in E,F,G
                foreach (var assetItem in balList.AssetItem.Assets)
                {
                    var rowAtIndx = TryGetAssetRow(assetRowIndx);
                    rowAtIndx.Append(
                    ConstructCell(assetItem.CatagoryItem, CellValues.String, 9, string.Format("{0}{1}", "E", assetRowIndx)),
                    ConstructCell(string.Empty, CellValues.String, 8, string.Format("{0}{1}", "F", assetRowIndx)),
                    ConstructCell(string.Empty, CellValues.String, 8, string.Format("{0}{1}", "G", assetRowIndx)));
                    assetCount++;
                    assetRowIndx++;
                    foreach (var item in assetItem.Items)
                    {
                        //Sub Catagory Names with amount
                        rowAtIndx = TryGetAssetRow(assetRowIndx);
                        rowAtIndx.Append(
                        ConstructCell(item.ItemName, CellValues.String, 8, string.Format("{0}{1}", "E", assetRowIndx)),
                        ConstructCell(item.Value, CellValues.String, 7, string.Format("{0}{1}", "F", assetRowIndx)),
                        ConstructCell(string.Empty, CellValues.String, 8, string.Format("{0}{1}", "G", assetRowIndx)));
                        assetCount++;
                        assetRowIndx++;
                    }
                    //total amount per catagory
                    rowAtIndx = TryGetAssetRow(assetRowIndx);
                    rowAtIndx.Append(
                    ConstructCell(string.Empty, CellValues.String, 8, string.Format("{0}{1}", "E", assetRowIndx)),
                    ConstructCell(string.Empty, CellValues.String, 8, string.Format("{0}{1}", "F", assetRowIndx)),
                    ConstructCell(assetItem.TotalValue, CellValues.String, 7, string.Format("{0}{1}", "G", assetRowIndx)));
                    assetCount++;
                    assetRowIndx++;
                }

                if (liabCount < assetCount)
                {
                    rowIndex = assetRowIndx;
                }
                else
                {
                    //need work out
                    for (UInt32 indx = assetRowIndx; indx <= rowIndex; indx++)
                    {
                        var rowEmtAsset = TryGetAssetRow(indx);
                        rowEmtAsset.Append(ConstructMergedCells(string.Empty, string.Format("{0}{1}", "E", indx), 3, 8));
                    }
                }

                //Add RowDictionary To sheet data
                foreach (var rowDic in RowDicSheetData)
                {
                    sheetData.AppendChild(rowDic.Value);
                }

                //empty row
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(
                    ConstructMergedCells(string.Empty, string.Format("{0}{1}", "A", rowIndex), 7, 8));
                sheetData.AppendChild(row);
                rowIndex++;

                //total summary row
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(
                    ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)),
                    ConstructCell(string.Empty, CellValues.String, 12, string.Format("{0}{1}", "B", rowIndex)),
                    ConstructCell(string.Empty, CellValues.String, 6, string.Format("{0}{1}", "C", rowIndex)),
                    ConstructCell(balList.LiabilityItem.TotalAmount, CellValues.String, 10, string.Format("{0}{1}", "D", rowIndex)),
                    ConstructCell(string.Empty, CellValues.String, 12, string.Format("{0}{1}", "E", rowIndex)),
                    ConstructCell(string.Empty, CellValues.String, 6, string.Format("{0}{1}", "F", rowIndex)),
                    ConstructCell(balList.AssetItem.TotalAmount, CellValues.String, 10, string.Format("{0}{1}", "G", rowIndex)));
                sheetData.AppendChild(row);
                rowIndex++;

                //disclimar row
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)));
                refCell = string.Format("{0}{1}", "B", rowIndex);
                row.Append(
                    ConstructMergedCells("As per our report of even date attached", refCell, 3, 13));
                mergeRef = IncrementXLColumn(refCell, 3);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });

                refCell = IncrementXLColumn(mergeRef);
                row.Append(
                    ConstructMergedCells(string.Empty, refCell, 3, 14));
                mergeRef = IncrementXLColumn(refCell, 3);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });

                sheetData.AppendChild(row);
                rowIndex++;

                //disclimar row
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)));
                refCell = string.Format("{0}{1}", "B", rowIndex);
                row.Append(
                    ConstructMergedCells(string.Empty, refCell, 3, 13));
                mergeRef = IncrementXLColumn(refCell, 3);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });

                refCell = IncrementXLColumn(mergeRef);
                row.Append(
                    ConstructMergedCells("For and on behalf of the Board", refCell, 3, 14));
                mergeRef = IncrementXLColumn(refCell, 3);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });

                sheetData.AppendChild(row);
                rowIndex++;

                //empty row
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)));
                refCell = string.Format("{0}{1}", "B", rowIndex);
                row.Append(
                    ConstructMergedCells(string.Empty, refCell, 3, 13));
                mergeRef = IncrementXLColumn(refCell, 3);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });

                refCell = IncrementXLColumn(mergeRef);
                row.Append(
                    ConstructMergedCells(string.Empty, refCell, 3, 14));
                mergeRef = IncrementXLColumn(refCell, 3);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });
                sheetData.AppendChild(row);
                rowIndex++;

                //empty row
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)));
                refCell = string.Format("{0}{1}", "B", rowIndex);
                row.Append(
                    ConstructMergedCells(string.Empty, refCell, 3, 13));
                mergeRef = IncrementXLColumn(refCell, 3);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });

                refCell = IncrementXLColumn(mergeRef);
                row.Append(
                    ConstructMergedCells(string.Empty, refCell, 3, 14));
                mergeRef = IncrementXLColumn(refCell, 3);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });
                sheetData.AppendChild(row);
                rowIndex++;

                //empty row
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)));
                refCell = string.Format("{0}{1}", "B", rowIndex);
                row.Append(
                    ConstructMergedCells(string.Format("Place : {0}", balList.CondoCity), refCell, 3, 13));
                mergeRef = IncrementXLColumn(refCell, 3);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });

                refCell = IncrementXLColumn(mergeRef);
                row.Append(
                    ConstructMergedCells(string.Empty, refCell, 3, 14));
                mergeRef = IncrementXLColumn(refCell, 3);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });
                sheetData.AppendChild(row);
                rowIndex++;

                //withness row
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)));
                refCell = string.Format("{0}{1}", "B", rowIndex);
                row.Append(
                    ConstructMergedCells(string.Format("Date : {0}", DateTime.Today.ToString("dd/MM/yyyy")), refCell, 3, 13));
                mergeRef = IncrementXLColumn(refCell, 3);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });
                refCell = IncrementXLColumn(mergeRef);
                row.Append(ConstructCell("President", CellValues.String, 11, refCell));
                refCell = IncrementXLColumn(refCell);
                row.Append(
                   ConstructMergedCells("Vice President", refCell, 2, 14));
                mergeRef = IncrementXLColumn(refCell, 2);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });
                sheetData.AppendChild(row);
                rowIndex++;

                //last row
                row = new Row();
                row.RowIndex = rowIndex;
                row.Append(ConstructCell(string.Empty, CellValues.String, 0, string.Format("{0}{1}", "A", rowIndex)));
                refCell = string.Format("{0}{1}", "B", rowIndex);
                row.Append(
                    ConstructMergedCells(string.Empty, refCell, 6, 12));
                mergeRef = IncrementXLColumn(refCell, 6);
                mergeCells.Append(new MergeCell() { Reference = new StringValue(string.Format("{0}:{1}", refCell, mergeRef)) });
                sheetData.AppendChild(row);
                rowIndex++;

                worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
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
                    new FontSize() { Val = 12 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }

                ),
                new Font( // Index 2 - group row
                    new FontSize() { Val = 11 }

                ),
                new Font( // Index 3 - bold-black-times
                    new FontSize() { Val = 12 },
                    new Bold(),
                    new FontName() { Val = "Times New Roman" }
                ),
                new Font( // Index 4 - bold-black-times-untedline
                    new FontSize() { Val = 12 },
                    new Bold(),
                    new FontName() { Val = "Times New Roman" },
                    new Underline() { Val = UnderlineValues.Single }
                ),
                new Font( // Index 5 - times
                    new FontSize() { Val = 12 },
                    new FontName() { Val = "Times New Roman" }
                ));

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } }) { PatternType = PatternValues.Solid }), // Index 2 - header
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "d3d3d3" } }) { PatternType = PatternValues.Solid }) // Index 3 - row group
                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder()),
                    new Border( // index 2 left only
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder()),
                    new Border( // index 3 right only
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder()),
                    new Border( // index 4 left-right 
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder()),
                    new Border( // index 5 left-right-bottom
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), //index 0 default
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, //index 1 body, defualt Left Text Align
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center } }, //index 2 header
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Right } }, //index 3 Right Text Align
                    new CellFormat { FontId = 2, FillId = 3, BorderId = 1, ApplyBorder = true }, // index 4 row group, defualt Left Text Align
                    new CellFormat { FontId = 2, FillId = 3, BorderId = 1, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Right } }, //index 5 row group Right Text Align

                    new CellFormat { FontId = 3, FillId = 0, BorderId = 1, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center } }, // index 6 bold center times
                    new CellFormat { FontId = 5, FillId = 0, BorderId = 4, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Right } }, // index 7 Right Text Align amount
                    new CellFormat { FontId = 5, FillId = 0, BorderId = 4, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left } }, // index 8 left Text Align items
                    new CellFormat { FontId = 4, FillId = 0, BorderId = 4, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left } }, // index 9 left Text Align withh bold-underline
                    new CellFormat { FontId = 4, FillId = 0, BorderId = 1, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Right } }, // index 10 right Text Align withh bold-underline-allborder
                    new CellFormat { FontId = 5, FillId = 0, BorderId = 0, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left } }, // index 11 with no border
                    new CellFormat { FontId = 5, FillId = 0, BorderId = 5, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left } }, // index 12 with U border
                    new CellFormat { FontId = 5, FillId = 0, BorderId = 2, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left } }, // index 13 left border only
                    new CellFormat { FontId = 5, FillId = 0, BorderId = 3, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left } } // index 14 right border only
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

        private Row TryGetAssetRow(UInt32 index)
        {
            Row retRow = null;
            if (RowDicSheetData.ContainsKey(index))
            {
                retRow = RowDicSheetData[index];
            }
            else
            {
                retRow = new Row();
                retRow.RowIndex = index;
                retRow.Append(
                    ConstructMergedCells(string.Empty, string.Format("{0}{1}", "A", index), 4, 8));
                RowDicSheetData.Add(index, retRow);
            }
            return retRow;
        }
    }
}
