using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Provalley.DigiCo.Reports.PDF
{
    public class ITextPageEvents : PdfPageEventHelper
    {
        // This is the contentbyte object of the writer
        PdfContentByte cb;
        // we will put the final number of pages in a template
        PdfTemplate template;
        // this is the BaseFont we are going to use for the header / footer
        BaseFont bf = null;
        // This keeps track of the creation time
        DateTime PrintTime = DateTime.Now;

        #region Properties
        private string _HeaderReportName;
        public string HeaderReportName
        {
            get { return _HeaderReportName; }
            set { _HeaderReportName = value; }
        }
        private string _HeaderCondoName;
        public string HeaderCondoName
        {
            get { return _HeaderCondoName; }
            set { _HeaderCondoName = value; }
        }
        private string _HeaderUserName;
        public string HeaderUserName
        {
            get { return _HeaderUserName; }
            set { _HeaderUserName = value; }
        }
        private Font _HeaderFont;
        public Font HeaderFont
        {
            get { return _HeaderFont; }
            set { _HeaderFont = value; }
        }
        private Font _FooterFont;
        public Font FooterFont
        {
            get { return _FooterFont; }
            set { _FooterFont = value; }
        }
        #endregion
        // we override the onOpenDocument method
        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            try
            {
                PrintTime = DateTime.Now;
                bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb = writer.DirectContent;
                template = cb.CreateTemplate(50, 50);
            }
            catch (DocumentException de)
            {
            }
            catch (System.IO.IOException ioe)
            {
            }
        }

        public override void OnStartPage(PdfWriter writer, Document document)
        {
            base.OnStartPage(writer, document);
            Rectangle pageSize = document.PageSize;
            if (!string.IsNullOrWhiteSpace(HeaderCondoName + HeaderReportName + HeaderUserName))
            {
                PdfPTable HeaderTable = new PdfPTable(2);
                HeaderTable.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                HeaderTable.DefaultCell.Border =Rectangle.NO_BORDER;
                HeaderTable.TotalWidth = pageSize.Width - 50;

                PdfPCell HeaderLeftCell = new PdfPCell(new Phrase(8, HeaderCondoName, HeaderFont));
                HeaderLeftCell.Padding = 5;
                HeaderLeftCell.PaddingBottom = 0;
                HeaderLeftCell.Border = Rectangle.NO_BORDER;
                HeaderTable.AddCell(HeaderLeftCell);

                PdfPCell HeaderRightCell = new PdfPCell(new Phrase(8, ("Date: " + PrintTime.ToString("dd-MM-yyyy")), HeaderFont));
                HeaderRightCell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                HeaderRightCell.Padding = 5;
                HeaderRightCell.PaddingBottom = 0;
                HeaderRightCell.Border = Rectangle.NO_BORDER;
                HeaderTable.AddCell(HeaderRightCell);

                HeaderLeftCell = new PdfPCell(new Phrase(8, HeaderReportName, HeaderFont));
                HeaderLeftCell.Padding = 5;
                HeaderLeftCell.PaddingBottom = 8;
                HeaderLeftCell.Border = Rectangle.NO_BORDER;
                HeaderTable.AddCell(HeaderLeftCell);

                HeaderRightCell = new PdfPCell(new Phrase(8, ("User : " + HeaderUserName), HeaderFont));
                HeaderRightCell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                HeaderRightCell.Padding = 5;
                HeaderRightCell.PaddingBottom = 8;
                HeaderRightCell.Border = Rectangle.NO_BORDER;
                HeaderTable.AddCell(HeaderRightCell);

                cb.SetRGBColorFill(0, 0, 0);
                HeaderTable.WriteSelectedRows(0, -1, pageSize.GetLeft(25), pageSize.GetTop(30), cb);
            }
        }
        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);
            int pageN = writer.PageNumber;
            String text = "Page " + pageN + " of ";
            float len = bf.GetWidthPoint(text, 10);
            Rectangle pageSize = document.PageSize;
            cb.SetRGBColorFill(100, 100, 100);
            cb.BeginText();
            cb.SetFontAndSize(bf, 10);
            cb.SetTextMatrix(pageSize.GetLeft(25), pageSize.GetBottom(30));
            cb.ShowText(text);
            cb.EndText();
            cb.AddTemplate(template, pageSize.GetLeft(25) + len, pageSize.GetBottom(30));

            cb.BeginText();
            cb.SetFontAndSize(bf, 10);
            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT,
            "Printed On " + PrintTime.ToString(),
            pageSize.GetRight(25),
            pageSize.GetBottom(30), 0);
            cb.EndText();
        }
        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);
            template.BeginText();
            template.SetFontAndSize(bf, 10);
            template.SetTextMatrix(0, 0);
            template.ShowText("" + writer.PageNumber);
            template.EndText();
        }
    }
}
