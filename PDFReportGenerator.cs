using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Provalley.DigiCo.Reports.PDF
{
    public class PDFReportGenerator
    {
        #region Properties
        public string CondoName { get; set; }
        public string ReportName { get; set; }
        public string CreatedBy { get; set; }
        public Dictionary<string, string> FilterData { get; set; }
        public List<float> ColWidth { get; set; }
        public DataTable FooterData { get; set; }
        #endregion

        public PDFReportGenerator(string reportName, string createdBy, string condoName)
        {
            CondoName = condoName;
            ReportName = reportName;
            CreatedBy = createdBy;
            FilterData = new Dictionary<string, string>();
            ColWidth = new List<float>();
            FooterData = new DataTable();
        }
        public byte[] CreatePDF(DataTable dataTable, Rectangle pageSize = null)
        {
            byte[] result = null;
            int _dtColCount = dataTable.Columns.Count;
            int _colWidth = ColWidth.Count;
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
            {
                using (Document document = new Document(pageSize ?? PageSize.A4, 25f, 25f, 80f, 40f))
                {
                    PdfWriter writer = PdfWriter.GetInstance(document, ms);
                    // Our custom Header and Footer is done using Event Handler
                    ITextPageEvents PageEventHandler = new ITextPageEvents();
                    // Define the page header
                    PageEventHandler.HeaderFont = FontFactory.GetFont(BaseFont.HELVETICA, 10, Font.NORMAL);
                    PageEventHandler.HeaderCondoName = CondoName;
                    PageEventHandler.HeaderReportName = ReportName;
                    PageEventHandler.HeaderUserName = CreatedBy;
                    writer.PageEvent = PageEventHandler;
                    document.Open();

                    if (FilterData.Any())
                    {
                        PdfPTable tableFilter = new PdfPTable(2);
                        tableFilter.WidthPercentage = 100;
                        tableFilter.SetWidths(new float[2] { 20f, 80f });
                        tableFilter.DefaultCell.Border = Rectangle.NO_BORDER;

                        PdfPCell cellLabel = new PdfPCell(new Phrase("Filters Applied :", FontFactory.GetFont(BaseFont.HELVETICA, 10, Font.BOLD)));

                        cellLabel.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        cellLabel.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cellLabel.Border = Rectangle.NO_BORDER;
                        cellLabel.Colspan = 2;
                        tableFilter.AddCell(cellLabel);
                        foreach (var item in FilterData)
                        {
                            PdfPCell cellFilterName = new PdfPCell(new Phrase(item.Key, FontFactory.GetFont(BaseFont.HELVETICA, 8, Font.NORMAL)));
                            cellFilterName.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                            cellFilterName.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                            cellFilterName.Border = Rectangle.NO_BORDER;
                            tableFilter.AddCell(cellFilterName);

                            PdfPCell cellFilterValue = new PdfPCell(new Phrase(item.Value, FontFactory.GetFont(BaseFont.HELVETICA, 8, Font.NORMAL)));
                            cellFilterValue.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                            cellFilterValue.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                            cellFilterValue.Border = Rectangle.NO_BORDER;
                            tableFilter.AddCell(cellFilterValue);

                        }
                        PdfPCell cellEmpty = new PdfPCell(new Phrase("\n", FontFactory.GetFont(BaseFont.HELVETICA, 8, Font.NORMAL)));
                        cellEmpty.Border = Rectangle.NO_BORDER;
                        cellEmpty.Colspan = 2;
                        cellEmpty.Padding = 8;
                        tableFilter.AddCell(cellEmpty);
                        document.Add(tableFilter);
                    }

                    PdfPTable tableData = new PdfPTable(_dtColCount);
                    tableData.WidthPercentage = 100;
                    if (_dtColCount == _colWidth)
                    {
                        tableData.SetWidths(ColWidth.ToArray());
                    }

                    //Set columns names in the pdf file
                    for (int k = 0; k < dataTable.Columns.Count; k++)
                    {
                        PdfPCell cell = new PdfPCell(new Phrase(dataTable.Columns[k].ColumnName, FontFactory.GetFont(BaseFont.HELVETICA, 10, Font.BOLD, BaseColor.WHITE)));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.BackgroundColor = BaseColor.GRAY;
                        tableData.AddCell(cell);
                    }
                    tableData.HeaderRows = 1;

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            PdfPCell cell = null;
                            if (dataTable.Columns[j].DataType == typeof(System.DateTime) && !(string.IsNullOrEmpty(dataTable.Rows[i][j].ToString())))
                            {
                                cell = new PdfPCell(new Phrase(((DateTime)dataTable.Rows[i][j]).ToString("dd/MM/yyyy"), FontFactory.GetFont(BaseFont.HELVETICA, 8, Font.NORMAL)));
                                cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                            }
                            else if (dataTable.Columns[j].ColumnName.ToUpper().Contains("RM"))
                            {

                                cell = new PdfPCell(new Phrase(String.Format("{0:N2}", dataTable.Rows[i][j]), FontFactory.GetFont(BaseFont.HELVETICA, 8, Font.NORMAL)));
                                cell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                            }
                            else
                            {
                                cell = new PdfPCell(new Phrase(dataTable.Rows[i][j].ToString(), FontFactory.GetFont(BaseFont.HELVETICA, 8, Font.NORMAL)));
                                cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                            }
                            cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                            tableData.AddCell(cell);
                        }
                    }
                    //Adding footer
                    if (FooterData.Rows.Count > 0)
                    {
                        var colSpan = _dtColCount - 2; //need to look into this count
                        for (int i = 0; i < FooterData.Rows.Count; i++)
                        {
                            for (int j = 0; j < FooterData.Columns.Count; j++)
                            {
                                PdfPCell cell = null;
                                decimal parseVal = 0;
                                if (decimal.TryParse(FooterData.Rows[i][j].ToString(), out parseVal))
                                {
                                    cell = new PdfPCell(new Phrase(String.Format("{0:N2}", FooterData.Rows[i][j]), FontFactory.GetFont(BaseFont.HELVETICA, 9, Font.NORMAL, BaseColor.BLACK)));
                                    cell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                                }
                                else
                                {
                                    cell = new PdfPCell(new Phrase(FooterData.Rows[i][j].ToString(), FontFactory.GetFont(BaseFont.HELVETICA, 9, Font.NORMAL, BaseColor.BLACK)));
                                    cell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                                }
                                if (j == 0)
                                {
                                    cell.Colspan = colSpan;
                                }
                                cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                                cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                                tableData.AddCell(cell);
                            }
                        }
                    }

                    document.Add(tableData);
                    document.Close();
                    result = ms.ToArray();
                }
            }

            return result;
        }

        public byte[] CreateRowGroupPDF(DataTable dataTable, string groupingColumn)
        {
            byte[] result = null;
            int _dtColCount = string.IsNullOrEmpty(groupingColumn) ? dataTable.Columns.Count : dataTable.Columns.Count - 1;
            int _colWidth = ColWidth.Count;
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
            {
                using (Document document = new Document(PageSize.A4, 25f, 25f, 80f, 40f))
                {
                    PdfWriter writer = PdfWriter.GetInstance(document, ms);
                    // Our custom Header and Footer is done using Event Handler
                    ITextPageEvents PageEventHandler = new ITextPageEvents();
                    // Define the page header
                    PageEventHandler.HeaderFont = FontFactory.GetFont(BaseFont.HELVETICA, 10, Font.NORMAL);
                    PageEventHandler.HeaderCondoName = CondoName;
                    PageEventHandler.HeaderReportName = ReportName;
                    PageEventHandler.HeaderUserName = CreatedBy;
                    writer.PageEvent = PageEventHandler;
                    document.Open();

                    if (FilterData.Any())
                    {
                        PdfPTable tableFilter = new PdfPTable(2);
                        tableFilter.WidthPercentage = 100;
                        tableFilter.SetWidths(new float[2] { 20f, 80f });
                        tableFilter.DefaultCell.Border = Rectangle.NO_BORDER;

                        // Fileter applied  lables
                        /* PdfPCell cellLabel = new PdfPCell(new Phrase("Filters Applied :", FontFactory.GetFont(BaseFont.HELVETICA, 10, Font.BOLD)));

                        cellLabel.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        cellLabel.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cellLabel.Border = Rectangle.NO_BORDER;
                        cellLabel.Colspan = 2;
                        tableFilter.AddCell(cellLabel); */
                        foreach (var item in FilterData)
                        {
                            PdfPCell cellFilterName = new PdfPCell(new Phrase(item.Key, FontFactory.GetFont(BaseFont.HELVETICA, 8, Font.NORMAL)));
                            cellFilterName.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                            cellFilterName.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                            cellFilterName.Border = Rectangle.NO_BORDER;
                            tableFilter.AddCell(cellFilterName);

                            PdfPCell cellFilterValue = new PdfPCell(new Phrase(item.Value, FontFactory.GetFont(BaseFont.HELVETICA, 8, Font.NORMAL)));
                            cellFilterValue.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                            cellFilterValue.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                            cellFilterValue.Border = Rectangle.NO_BORDER;
                            tableFilter.AddCell(cellFilterValue);

                        }
                        PdfPCell cellEmpty = new PdfPCell(new Phrase("\n", FontFactory.GetFont(BaseFont.HELVETICA, 8, Font.NORMAL)));
                        cellEmpty.Border = Rectangle.NO_BORDER;
                        cellEmpty.Colspan = 2;
                        cellEmpty.Padding = 8;
                        tableFilter.AddCell(cellEmpty);
                        document.Add(tableFilter);
                    }

                    PdfPTable tableData = new PdfPTable(_dtColCount);
                    tableData.WidthPercentage = 100;
                    if (_dtColCount == _colWidth)
                    {
                        tableData.SetWidths(ColWidth.ToArray());
                    }

                    //Set columns names in the pdf file
                    for (int k = 0; k < dataTable.Columns.Count; k++)
                    {
                        if (groupingColumn.Equals(dataTable.Columns[k].ColumnName, StringComparison.InvariantCultureIgnoreCase))
                        {
                            continue;
                        }
                        PdfPCell cell = new PdfPCell(new Phrase(dataTable.Columns[k].ColumnName, FontFactory.GetFont(BaseFont.HELVETICA, 10, Font.BOLD, BaseColor.WHITE)));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.BackgroundColor = BaseColor.DARK_GRAY;
                        tableData.AddCell(cell);
                    }
                    tableData.HeaderRows = 1;

                    var groupRowList = (from row in dataTable.AsEnumerable()
                                        select row[groupingColumn].ToString()).Distinct().ToList();
                    foreach (var groupName in groupRowList)
                    {
                        var groupedRows = (from row in dataTable.AsEnumerable()
                                           where row.Field<string>(groupingColumn) == groupName
                                           select row);
                        var sumOfDr = groupedRows.Sum(s => s.Field<decimal>("Debit(RM)"));
                        var sumOfCr = groupedRows.Sum(s => s.Field<decimal>("Credit(RM)"));
                        var balSum = sumOfDr - sumOfCr;
                        var strBalSum = balSum < 0 ? string.Format("({0:N2})", balSum * -1) : string.Format("{0:N2}", balSum);

                        PdfPCell cell = new PdfPCell(new Phrase(groupName, FontFactory.GetFont(BaseFont.HELVETICA, 9, Font.NORMAL)));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                        cell.Colspan = _dtColCount - 1;
                        tableData.AddCell(cell);

                        cell = new PdfPCell(new Phrase(strBalSum, FontFactory.GetFont(BaseFont.HELVETICA, 9, Font.NORMAL)));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                        tableData.AddCell(cell);

                        foreach (var row in groupedRows)
                        {
                            foreach (DataColumn col in dataTable.Columns)
                            {
                                if (!col.ColumnName.Equals(groupingColumn, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    if (col.ColumnName == "Debit(RM)" || col.ColumnName == "Credit(RM)")
                                    {
                                        cell = new PdfPCell(new Phrase(string.Format("{0:N2}", row.Field<decimal>(col)), FontFactory.GetFont(BaseFont.HELVETICA, 8, Font.NORMAL)));
                                        cell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                                    }
                                    else
                                    {
                                        cell = new PdfPCell(new Phrase(row[col.ColumnName].ToString(), FontFactory.GetFont(BaseFont.HELVETICA, 8, Font.NORMAL)));
                                        cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                                    }
                                    tableData.AddCell(cell);
                                }
                            }
                        }
                    }

                    document.Add(tableData);
                    document.Close();
                    result = ms.ToArray();
                }
            }

            return result;
        }
    }
}
