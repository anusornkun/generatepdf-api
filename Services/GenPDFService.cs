using iTextGenPDF.Api.Models;
using iTextGenPDF.Api.Services.Interface;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;

namespace iTextGenPDF.Api.Services
{
    public class GenPDFService : IGenPDFService
    {
        #region CreatePDF Credit_Note/Debit_Note
        public FileResponseDataBinding CreateCreditDebitNote(Refund creditNoteModel)
        {
            var response = new FileResponseDataBinding();
            PdfPTable table = null;
            PdfWriter writer = null;

            try
            {
                Document document = new Document(PageSize.A4, 0f, 0f, 300f, 25f);

                using (MemoryStream ms = new MemoryStream())
                {
                    writer = PdfWriter.GetInstance(document, ms);
                    string fontpath_normal = Path.Combine(Environment.CurrentDirectory, "Fonts", "THSarabunNew.ttf");
                    string fontpathBold = Path.Combine(Environment.CurrentDirectory, "Fonts", "THSarabunNew Bold.ttf");
                    BaseFont EnCodefont = BaseFont.CreateFont(fontpath_normal, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    BaseFont EnCodefontBold = BaseFont.CreateFont(fontpathBold, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    Font forHeaders = new Font(EnCodefontBold, 14, Font.NORMAL);
                    Font forDetails = new Font(EnCodefont, 14, Font.NORMAL);
                    //Font forNote = new Font(EnCodefont, 10, Font.NORMAL);
                    writer.PageEvent = new HeaderCreditDebitNote() { queryModel = creditNoteModel };

                    document.Open();

                    Image image = Image.GetInstance(Environment.CurrentDirectory + "/Images/csp.jpg");
                    image.ScaleAbsolute(102, 30);

                    table = new PdfPTable(10);
                    table.TotalWidth = 600f;
                    table.LockedWidth = true;
                    table.SetWidths(new float[] { 30f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 20f });
                    table.DefaultCell.Border = Rectangle.NO_BORDER;
                    table.DefaultCell.NoWrap = true;
                    table.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT;

                    //details start--------------------------------------------------------
                    for (int i = 1; i <= 13; i++)
                    {
                        Random rnd = new Random();
                        double rate = (0.85 * rnd.Next(2, 17));
                        table.AddCell(new PdfPCell(new Phrase($"{i}. ", forDetails)) { HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 });
                        table.AddCell(new PdfPCell(new Phrase($"ค่าบริการส่ง E-Statement  Credit Card by E-Mail = > END User ประจำเดือน มิถุนายน 2563ประจำเดือนมิถุนายน2563ประจำเดือนมิถุนายน2563ประจำเดือนมิถุนายน2563ประจำเดือนมิถุนายน2563", forDetails)) { HorizontalAlignment = Element.ALIGN_LEFT, Border = 0, Colspan = 5 });
                        table.AddCell(new PdfPCell(new Phrase("18,173.00", forDetails)) { HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 });
                        table.AddCell(new PdfPCell(new Phrase($"{Math.Round(rate, 2).ToString("F2")}", forDetails)) { HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 });
                        table.AddCell(new PdfPCell(new Phrase("15,447.05", forDetails)) { HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 });
                        table.AddCell("");
                    }
                    //details stop------------------------------------------------------------
                    document.Add(table);

                    var pointer = writer.GetVerticalPosition(true);
                    if (pointer < 240)
                    {
                        document.NewPage();
                    }
                    table = new PdfPTable(3);
                    table.TotalWidth = 600f;
                    table.LockedWidth = true;
                    table.SetWidths(new float[] { 200f, 300f, 100f });
                    //--------------------------
                    table.AddCell(new PdfPCell(new Phrase("มูลค่าของสินค้าหรือบริการตามใบเสร็จรับเงิน/ใบกำากับภาษีเดิม", forHeaders))
                    {
                        HorizontalAlignment = Element.ALIGN_RIGHT,
                        Colspan = 2,
                        Border = 0,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f,
                        FixedHeight = 25f
                    });
                    table.AddCell(new PdfPCell(new Phrase("132,773.50", forDetails))
                    {
                        HorizontalAlignment = Element.ALIGN_RIGHT,
                        Border = 0,
                        PaddingRight = 25f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f
                    });
                    //--------------------------
                    table.AddCell(new PdfPCell(new Phrase("มูลค่าของสินค้าหรือบริการตามใบเสร็จรับเงิน/ใบกำากับภาษีใหม่", forHeaders))
                    {
                        HorizontalAlignment = Element.ALIGN_RIGHT,
                        Colspan = 2,
                        Border = 0,
                        FixedHeight = 25f
                    });
                    table.AddCell(new PdfPCell(new Phrase("0.00", forDetails))
                    {
                        HorizontalAlignment = Element.ALIGN_RIGHT,
                        Border = 0,
                        PaddingRight = 25f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f,
                    });
                    //--------------------------
                    table.AddCell(new PdfPCell(new Phrase("ผลต่าง", forHeaders))
                    {
                        HorizontalAlignment = Element.ALIGN_RIGHT,
                        Colspan = 2,
                        Border = 0,
                        FixedHeight = 25f
                    });
                    table.AddCell(new PdfPCell(new Phrase("132,773.50", forDetails))
                    {
                        HorizontalAlignment = Element.ALIGN_RIGHT,
                        Border = 0,
                        PaddingRight = 25f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f,
                    });
                    //--------------------------
                    table.AddCell(new PdfPCell(new Phrase("ภาษีมูลค่าเพิ่ม", forHeaders))
                    {
                        HorizontalAlignment = Element.ALIGN_RIGHT,
                        Colspan = 2,
                        Border = 0,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f,
                        FixedHeight = 25f
                    });
                    table.AddCell(new PdfPCell(new Phrase("21.56", forDetails))
                    {
                        HorizontalAlignment = Element.ALIGN_RIGHT,
                        Border = 0,
                        PaddingRight = 25f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f
                    });
                    //--------------------------
                    table.AddCell(new PdfPCell(new Phrase("(เก้าพันสองร้อยเก้าสิบสี่บาทสิบห้าสตางค์)", forDetails))
                    {
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        Border = 0,
                        PaddingLeft = 25f,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f,
                        FixedHeight = 25f
                    });
                    table.AddCell(new PdfPCell(new Phrase("รวมเป็นเงินทั้งสิ้น", forHeaders))
                    {
                        HorizontalAlignment = Element.ALIGN_RIGHT,
                        Border = 0,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f
                    });
                    table.AddCell(new PdfPCell(new Phrase("9,294.15", forDetails))
                    {
                        HorizontalAlignment = Element.ALIGN_RIGHT,
                        Border = 0,
                        PaddingRight = 25f,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f
                    });
                    //--------------------------
                    Phrase phrase = new Phrase();
                    if (creditNoteModel.TypeCode == "81")
                    {
                        phrase.Add(new Chunk($"เหตุผลของการลดหนี้ : CDNG02\n", forHeaders));

                    }
                    if (creditNoteModel.TypeCode == "80")
                    {
                        phrase.Add(new Chunk($"เหตุผลของการเพิ่มหนี้ : DBNG02\n", forHeaders));
                    }
                    phrase.Add(new Chunk($"            สินค้าชารุดเสีย", forDetails));
                    table.AddCell(new PdfPCell(phrase)
                    {
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        PaddingLeft = 25f,
                        Border = 0,
                        Colspan = 3
                    });
                    //--------------------------
                    table.AddCell(new PdfPCell(new Phrase("หมายเหตุ : เอกสารนี้ได้จัดทำาและส่งข้อมูลให้แก่กรมสรรพากรด้วยวิธีการทางอิเล็กทรอนิกส์", forHeaders))
                    {
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        PaddingLeft = 25f,
                        Border = 0,
                        Colspan = 3,
                        FixedHeight = 25f
                    });

                    table.WriteSelectedRows(0, 8, 0, document.Bottom + 200, writer.DirectContent);

                    document.Close();
                    response.File = ms.ToArray();
                    response.StatusCode = 200;
                }
            }
            catch (Exception ex)
            {
                response.StatusCode = 500;
                response.Message = ex.Message;
                return response;
            }

            return response;
        }

        public partial class HeaderCreditDebitNote : PdfPageEventHelper
        {
            public Refund queryModel { get; set; }
            PdfTemplate template;
            PdfContentByte cb;

            public int TotalNumber { get; set; } = 0;

            public string fontpath_normal { get; set; } = Path.Combine(Environment.CurrentDirectory, "Fonts", "THSarabunNew.ttf");
            public string fontpathBold { get; set; } = Path.Combine(Environment.CurrentDirectory, "Fonts", "THSarabunNew Bold.ttf");

            public override void OnOpenDocument(PdfWriter writer, Document document)
            {
                cb = writer.DirectContent;
                template = cb.CreateTemplate(550, 50);
            }
            public override void OnStartPage(PdfWriter writer, Document document)
            {
                BaseFont EnCodefont = BaseFont.CreateFont(fontpath_normal, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                BaseFont EnCodefontBold = BaseFont.CreateFont(fontpathBold, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                base.OnStartPage(writer, document);
                int pageN = writer.PageNumber;
                string strPage = $"หน้าที่ {pageN.ToString("N0")} / ";
                float len = EnCodefont.GetWidthPoint(strPage, 11);

                Font forHeaders = new Font(EnCodefontBold, 14, Font.NORMAL);
                Font forDetails = new Font(EnCodefont, 14, Font.NORMAL);
                Font forNote = new Font(EnCodefont, 10, Font.NORMAL);


                Font forCSP = new Font(EnCodefontBold, 13, Font.NORMAL);
                Font forNameTH = new Font(EnCodefontBold, 19, Font.NORMAL);
                Font forHeadOffice = new Font(EnCodefontBold, 14, Font.NORMAL, BaseColor.BLACK);

                Image image = Image.GetInstance(Environment.CurrentDirectory + "/Images/csp.jpg");
                image.ScalePercent(18f);
                PdfPTable table = null;
                table = new PdfPTable(10);
                table.TotalWidth = 600f;
                table.LockedWidth = true;
                table.SetWidths(new float[] { 25f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 25f });
                table.DefaultCell.Border = Rectangle.NO_BORDER;
                table.DefaultCell.NoWrap = true;
                table.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT;
                table.HeaderRows = 9;

                //header start-----------------------------------------------------------
                table.AddCell(new PdfPCell(image) { VerticalAlignment = Element.ALIGN_BOTTOM, HorizontalAlignment = Element.ALIGN_RIGHT, Colspan = 3, Rowspan = 2, Border = 0 });
                table.AddCell(new PdfPCell(new Phrase("บริษัท จันวาณิชย์ ซีเคียวริตี้พริ้นท์ติ้ง จำกัด", forNameTH)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Colspan = 7, Border = 0 });
                //--------------------
                Phrase phrase = new Phrase();
                phrase.Add(new Chunk($"CHAN WANICH SECURITY PRINTING COMPANY LIMITED ", forHeadOffice));
                phrase.Add(new Chunk($"             เลขที่ 81-0000000002", forHeaders));
                table.AddCell(new PdfPCell(phrase)
                {
                    VerticalAlignment = Element.ALIGN_TOP,
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    Colspan = 7,
                    Border = 0
                });
                //--------------------
                table.AddCell(new PdfPCell(new Phrase("")) { VerticalAlignment = Element.ALIGN_BOTTOM, HorizontalAlignment = Element.ALIGN_RIGHT, Colspan = 4, Rowspan = 2, Border = 0 });
                table.AddCell(new PdfPCell(new Phrase("เลขประจำตัวผู้เสียภาษีอากร", forHeadOffice)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Colspan = 6, Border = 0 });         //--------------------
                table.AddCell(new PdfPCell(new Phrase("      0105533079571      00002 : พระประแดง", forHeadOffice)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Colspan = 6, Border = 0 });
                //--------------------
                table.AddCell("");
                table.AddCell(new PdfPCell(new Phrase("699 ถนนสีลม แขวงสีลม เขตบางรัก\nกรุงเทพมหานคร 10500\nโทร.(662) 635-3333 โทรสาร (662) 236-7176", forNote)) { VerticalAlignment = Element.ALIGN_BOTTOM, HorizontalAlignment = Element.ALIGN_LEFT, NoWrap = true, Border = 0, PaddingBottom = 10f });
                table.AddCell("");
                if (queryModel.TypeCode == "80")
                    table.AddCell(new PdfPCell(new Phrase("ใบเพิ่มหนี้     \n(Debit Note)     \n", forNameTH)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Colspan = 4, Border = 0, PaddingBottom = 8f });
                if (queryModel.TypeCode == "81")
                    table.AddCell(new PdfPCell(new Phrase("ใบลดหนี้     \n(Credit Note)     \n", forNameTH)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Colspan = 4, Border = 0, PaddingBottom = 8f });

                table.AddCell(new PdfPCell(new Phrase("699 SILOM ROAD, SILOM, BANGRAK\nBANGKOK 10500, THAILAND\nTEL.(662) 635-3333 FAX.(662) 236-7176", forNote)) { VerticalAlignment = Element.ALIGN_BOTTOM, HorizontalAlignment = Element.ALIGN_LEFT, NoWrap = true, Border = 0, PaddingBottom = 10f });
                table.AddCell("");
                table.AddCell("");
                //--------------------
                table.AddCell(new PdfPCell(new Phrase(""))
                {
                    Border = 0,
                    PaddingBottom = 8f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                phrase = new Phrase();
                phrase.Add(new Chunk($"รหัสลูกค้า ", forHeadOffice));
                phrase.Add(new Chunk($"004346", forDetails));
                table.AddCell(new PdfPCell(phrase)
                {
                    VerticalAlignment = Element.ALIGN_TOP,
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    NoWrap = true,
                    Colspan = 2,
                    Border = 0,
                    BorderColorTop = new BaseColor(System.Drawing.Color.Black),
                    BorderWidthTop = 0.1f
                });

                phrase = new Phrase();
                phrase.Add(new Chunk($"TAX ID : ", forHeadOffice));
                phrase.Add(new Chunk($"0107557000489 : สำนักงานใหญ่", forDetails));
                table.AddCell(new PdfPCell(phrase)
                {
                    VerticalAlignment = Element.ALIGN_TOP,
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    NoWrap = true,
                    Colspan = 3,
                    Border = 0,
                    BorderColorTop = new BaseColor(System.Drawing.Color.Black),
                    BorderWidthTop = 0.1f
                });
                phrase = new Phrase();
                phrase.Add(new Chunk($"วันที่ ", forHeadOffice));
                phrase.Add(new Chunk($"29 กุมภาพันธ์ 2563\n", forDetails));
                phrase.Add(new Chunk($"เลขที่ใบส่งของ/ใบแจ้งหนี้ ", forHeadOffice));
                phrase.Add(new Chunk($"2BT-DB1-1804-0022\n", forDetails));
                phrase.Add(new Chunk($"เลขที่ใบเสร็จรับเงิน/ใบกำากับภาษี ", forHeadOffice));
                phrase.Add(new Chunk($"2BI-DB1-1804-0030", forDetails));
                table.AddCell(new PdfPCell(phrase)
                {
                    VerticalAlignment = Element.ALIGN_TOP,
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    Colspan = 4,
                    Rowspan = 3,
                    Border = 0,
                    BorderColorTop = new BaseColor(System.Drawing.Color.Black),
                    BorderWidthTop = 0.1f
                });
                //--------------------
                table.AddCell("");
                table.AddCell(new PdfPCell(new Phrase("ชื่อลูกค้า/ที่อยู่", forHeaders)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, NoWrap = true, Colspan = 6, Border = 0 });
                //--------------------
                table.AddCell("");
                table.AddCell(new PdfPCell(new Phrase("บจ. จันวาณิชย์\n699 สีลม ตำาบลบางพลีใหญ่ อำาเภอบางพลี จังหวัดสมุทรปราการ 10540", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Colspan = 6, Border = 0 });
                //--------------------
                table.AddCell("");
                table.AddCell(new PdfPCell(new Phrase("บริษัทฯ ได้เดบิตบัญชีของท่านตามรายการต่อไปนี้", forHeaders)) { VerticalAlignment = Element.ALIGN_BOTTOM, HorizontalAlignment = Element.ALIGN_LEFT, Colspan = 5, Border = 0, PaddingBottom = 8f, FixedHeight = 50f });
                Phrase notePhrase = new Phrase();
                notePhrase.Add(new Chunk($"อัตราภาษีร้อยละ ", forHeaders));
                Chunk chunk = new Chunk(" 7.00");
                chunk.SetUnderline(0.5f, -1);
                chunk.Font = forHeaders;
                notePhrase.Add(chunk);
                notePhrase.Add(new Chunk($" อัตราภาษีศูนย์", forHeaders));
                table.AddCell(new PdfPCell(notePhrase)
                {
                    VerticalAlignment = Element.ALIGN_BOTTOM,
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    Colspan = 4,
                    Border = 0,
                    PaddingBottom = 8f,
                    FixedHeight = 50f
                });

                //--------------------
                table.AddCell(new PdfPCell(new Phrase(""))
                {
                    Border = 0,
                    PaddingBottom = 8f,
                    BorderColorBottom = BaseColor.BLACK,
                    BorderWidthBottom = 0.1f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                table.AddCell(new PdfPCell(new Phrase("รายการ", forHeaders))
                {
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Colspan = 5,
                    Border = 0,
                    PaddingBottom = 8f,
                    BorderColorBottom = BaseColor.BLACK,
                    BorderWidthBottom = 0.1f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                table.AddCell(new PdfPCell(new Phrase("จำนวน", forHeaders))
                {
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Border = 0,
                    NoWrap = true,
                    PaddingBottom = 8f,
                    BorderColorBottom = BaseColor.BLACK,
                    BorderWidthBottom = 0.1f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                table.AddCell(new PdfPCell(new Phrase("หน่วยละ", forHeaders))
                {
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Border = 0,
                    NoWrap = true,
                    PaddingBottom = 8f,
                    BorderColorBottom = BaseColor.BLACK,
                    BorderWidthBottom = 0.1f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                table.AddCell(new PdfPCell(new Phrase("จำนวนเงิน", forHeaders))
                {
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Border = 0,
                    NoWrap = true,
                    PaddingBottom = 8f,
                    BorderColorBottom = BaseColor.BLACK,
                    BorderWidthBottom = 0.1f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                table.AddCell(new PdfPCell(new Phrase(""))
                {
                    Border = 0,
                    PaddingBottom = 8f,
                    BorderColorBottom = BaseColor.BLACK,
                    BorderWidthBottom = 0.1f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                //----------------------
                table.AddCell(new PdfPCell(new Phrase("")) { Colspan = 10 });
                table.WriteSelectedRows(0, -1, document.LeftMargin, document.PageSize.Height - 30, writer.DirectContent);


                cb.BeginText();
                cb.SetFontAndSize(EnCodefont, 11);
                cb.SetTextMatrix(482 - len, document.PageSize.Height - 80);
                cb.ShowText(strPage);


                cb.EndText();
                cb.AddTemplate(template, 482, document.PageSize.Height - 80);

            }
            public override void OnCloseDocument(PdfWriter writer, Document document)
            {
                BaseFont EnCodefont = BaseFont.CreateFont(fontpath_normal, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                base.OnCloseDocument(writer, document);

                template.BeginText();
                template.SetFontAndSize(EnCodefont, 11);
                template.SetTextMatrix(0, 0);
                template.ShowText("" + (writer.PageNumber.ToString("N0")));
                template.EndText();

            }
        }
        #endregion

        #region CreatePDF Receipt/Tax Invoice
        public FileResponseDataBinding CreateTaxInvoice(Payment payment)
        {
            var response = new FileResponseDataBinding();
            PdfPTable table = null;
            PdfWriter writer = null;

            try
            {
                Document document = new Document(PageSize.A4, 0f, 0f, 340f, 25f);

                using (MemoryStream ms = new MemoryStream())
                {
                    writer = PdfWriter.GetInstance(document, ms);
                    string fontpath_normal = Path.Combine(Environment.CurrentDirectory, "Fonts", "THSarabunNew.ttf");
                    string fontpathBold = Path.Combine(Environment.CurrentDirectory, "Fonts", "THSarabunNew Bold.ttf");
                    BaseFont EnCodefont = BaseFont.CreateFont(fontpath_normal, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    BaseFont EnCodefontBold = BaseFont.CreateFont(fontpathBold, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    Font forHeaders = new Font(EnCodefontBold, 14, Font.NORMAL);
                    Font forDetails = new Font(EnCodefont, 14, Font.NORMAL);
                    Font forNote = new Font(EnCodefont, 10, Font.NORMAL);
                    writer.PageEvent = new Header() { queryModel = payment };

                    document.Open();

                    Image image = Image.GetInstance(Environment.CurrentDirectory + "/Images/csp.jpg");
                    image.ScaleAbsolute(102, 30);

                    table = new PdfPTable(10);
                    table.TotalWidth = 600f;
                    table.LockedWidth = true;
                    table.SetWidths(new float[] { 30f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 20f });
                    table.DefaultCell.Border = Rectangle.NO_BORDER;
                    table.DefaultCell.NoWrap = true;
                    table.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT;

                    //details start--------------------------------------------------------
                    for (int i = 1; i <= 11; i++)
                    {
                        Random rnd = new Random();
                        double rate = (0.85 * rnd.Next(2, 17));
                        table.AddCell(new PdfPCell(new Phrase($"{i}. ", forDetails)) { HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 });
                        table.AddCell(new PdfPCell(new Phrase($"ค่าบริการส่ง E-Statement  Credit Card by E-Mail = > END User ประจำเดือน มิถุนายน 2563ประจำเดือนมิถุนายน2563ประจำเดือนมิถุนายน2563ประจำเดือนมิถุนายน2563ประจำเดือนมิถุนายน2563", forDetails)) { HorizontalAlignment = Element.ALIGN_LEFT, Border = 0, Colspan = 5 });
                        table.AddCell(new PdfPCell(new Phrase("18,173.00", forDetails)) { HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 });
                        table.AddCell(new PdfPCell(new Phrase($"{Math.Round(rate, 2).ToString("F2")}", forDetails)) { HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 });
                        table.AddCell(new PdfPCell(new Phrase("15,447.05", forDetails)) { HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 });
                        table.AddCell("");
                    }
                    //details stop------------------------------------------------------------
                    document.Add(table);

                    var pointer = writer.GetVerticalPosition(true);
                    if (pointer < 216)
                    {
                        document.NewPage();
                    }
                    table = new PdfPTable(4);
                    table.TotalWidth = 600f;
                    table.LockedWidth = true;
                    table.SetWidths(new float[] { 50f, 350f, 100f, 100f });

                    table.AddCell(new PdfPCell(new Phrase("บริษัทฯ ได้รับการส่งเสริมการลงทุน (BOI) โปรดยกเว้นการหักภาษี ณ ที่จ่าย", forHeaders))
                    {
                        VerticalAlignment = Element.ALIGN_MIDDLE,
                        HorizontalAlignment = Element.ALIGN_CENTER,
                        Colspan = 2,
                        Rowspan = 2,
                        Border = 0,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f
                    });
                    table.AddCell(new PdfPCell(new Phrase("รวมมูลค่า\nAmount", forHeaders))
                    {
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        Border = 0,
                        PaddingBottom = 8f,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f,
                        BorderColorLeft = BaseColor.BLACK,
                        BorderWidthLeft = 0.1f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f
                    });
                    table.AddCell(new PdfPCell(new Phrase("23,624.90", forDetails))
                    {
                        VerticalAlignment = Element.ALIGN_MIDDLE,
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        Border = 0,
                        PaddingBottom = 8f,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f
                    });
                    table.AddCell(new PdfPCell(new Phrase("ภาษีมูลค่าเพิ่ม\nVat 7%", forHeaders))
                    {
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        Border = 0,
                        PaddingBottom = 8f,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f,
                        BorderColorLeft = BaseColor.BLACK,
                        BorderWidthLeft = 0.1f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f
                    });
                    table.AddCell(new PdfPCell(new Phrase("1,653.74", forDetails))
                    {
                        VerticalAlignment = Element.ALIGN_MIDDLE,
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        Border = 0,
                        PaddingBottom = 8f,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f
                    });
                    table.AddCell(new PdfPCell(new Phrase("บาท\nBaht", forHeaders))
                    {
                        PaddingLeft = 10f,
                        VerticalAlignment = Element.ALIGN_MIDDLE,
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        Border = 0,
                        PaddingBottom = 8f,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f
                    });
                    table.AddCell(new PdfPCell(new Phrase("สองหมื่นห้าพันสองร้อยเจ็ดสิบแปดบาทหกสิบสี่สตางค์", forDetails))
                    {
                        VerticalAlignment = Element.ALIGN_MIDDLE,
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        Border = 0,
                        PaddingBottom = 8f,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f
                    });
                    table.AddCell(new PdfPCell(new Phrase("รวมเงินทั้งสิ้น\nTotal Amount", forHeaders))
                    {
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        Border = 0,
                        PaddingBottom = 8f,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f,
                        BorderColorLeft = BaseColor.BLACK,
                        BorderWidthLeft = 0.1f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f
                    });
                    table.AddCell(new PdfPCell(new Phrase("25,278.64", forDetails))
                    {
                        VerticalAlignment = Element.ALIGN_MIDDLE,
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        Border = 0,
                        PaddingBottom = 8f,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f
                    });
                    table.AddCell(new PdfPCell(new Phrase("หมายเหตุ", forHeaders))
                    {
                        PaddingLeft = 10f,
                        VerticalAlignment = Element.ALIGN_MIDDLE,
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        Border = 0,
                    });
                    Phrase notePhrase = new Phrase();
                    notePhrase.Add(new Chunk($": เอกสารนี้ได้จัดทำและส่งข้อมูลให้แก่กรมสรรพากรด้วยวิธีการทางอิเล็กทรอนิกส์\n" +
                        $": ใบเสร็จรับเงินฉบับนี้จะสมบูรณ์ต่อเมื่อ บริษัทฯ ได้รับชำระเงินตามจำนวนเงินที่ระบุในใบเสร็จรับเงิน/ใบกำกับภาษี\n" +
                        $"  และ/หรือ สามารถเรียกเก็บเงินตามเช็คเรียบร้อยแล้ว กรณีชำระด้วยเช็ค ขีดคร่อมเฉพาะ ", forNote));
                    Chunk chunk = new Chunk(" A/C PAYEE ONLY ");
                    chunk.SetUnderline(0.5f, -1);
                    chunk.SetUnderline(0.5f, 6);
                    chunk.Font = forNote;
                    notePhrase.Add(chunk);
                    notePhrase.Add(new Chunk($" ระบุชื่อในนาม\n  \"บริษัท จันวาณิชย์ ซีเคียวริตี้ พริ้นท์ติ้ง จำกัด\" เท่านั้น\n", forNote));
                    notePhrase.Add(new Chunk($": บริษัทฯ จะคิดค่าเสียหาย 2 % ต่อเดือน ของยอดหนี้ค่าสินค้า เมื่อเกินกำหนดชำระ", forNote));
                    table.AddCell(new PdfPCell(notePhrase)
                    {
                        VerticalAlignment = Element.ALIGN_MIDDLE,
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        Border = 0,
                    });
                    table.AddCell(new PdfPCell(new Phrase("6CI-DN0-2007-0034", forDetails))
                    {
                        Colspan = 2,
                        VerticalAlignment = Element.ALIGN_MIDDLE,
                        HorizontalAlignment = Element.ALIGN_CENTER,
                        PaddingBottom = 8f,
                        BorderColorBottom = BaseColor.BLACK,
                        BorderWidthBottom = 0.1f,
                        BorderColorLeft = BaseColor.BLACK,
                        BorderWidthLeft = 0.1f,
                        BorderColorTop = BaseColor.BLACK,
                        BorderWidthTop = 0.1f
                    });

                    table.WriteSelectedRows(0, 4, 0, document.Bottom + 175, writer.DirectContent);

                    document.Close();
                    response.File = ms.ToArray();
                    response.StatusCode = 200;
                }
            }
            catch (Exception ex)
            {
                response.StatusCode = 500;
                response.Message = ex.Message;
                return response;
            }

            return response;
        }

        public partial class Header : PdfPageEventHelper
        {
            public Payment queryModel { get; set; }
            PdfTemplate template;
            PdfContentByte cb;

            public int TotalNumber { get; set; } = 0;

            public string fontpath_normal { get; set; } = Path.Combine(Environment.CurrentDirectory, "Fonts", "THSarabunNew.ttf");
            public string fontpathBold { get; set; } = Path.Combine(Environment.CurrentDirectory, "Fonts", "THSarabunNew Bold.ttf");

            public override void OnOpenDocument(PdfWriter writer, Document document)
            {
                cb = writer.DirectContent;
                template = cb.CreateTemplate(550, 50);
            }
            public override void OnStartPage(PdfWriter writer, Document document)
            {
                BaseFont EnCodefont = BaseFont.CreateFont(fontpath_normal, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                BaseFont EnCodefontBold = BaseFont.CreateFont(fontpathBold, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                base.OnStartPage(writer, document);
                int pageN = writer.PageNumber;
                string strPage = $"หน้าที่ {pageN.ToString("N0")} / ";
                float len = EnCodefont.GetWidthPoint(strPage, 11);

                Font forHeaders = new Font(EnCodefontBold, 14, Font.NORMAL);
                Font forDetails = new Font(EnCodefont, 14, Font.NORMAL);
                Font forNote = new Font(EnCodefont, 10, Font.NORMAL);


                Font forCSP = new Font(EnCodefontBold, 13, Font.NORMAL);
                Font forNameTH = new Font(EnCodefontBold, 19, Font.NORMAL);
                Font forHeadOffice = new Font(EnCodefontBold, 14, Font.NORMAL, BaseColor.BLACK);

                Image image = Image.GetInstance(Environment.CurrentDirectory + "/Images/csp.jpg");
                image.ScaleAbsolute(102, 30);
                PdfPTable table = null;
                table = new PdfPTable(10);
                table.TotalWidth = 600f;
                table.LockedWidth = true;
                table.SetWidths(new float[] { 25f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 68.75f, 25f });
                table.DefaultCell.Border = Rectangle.NO_BORDER;
                table.DefaultCell.NoWrap = true;
                table.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT;
                table.HeaderRows = 9;

                //header start-----------------------------------------------------------
                table.AddCell(new PdfPCell(image) { VerticalAlignment = Element.ALIGN_BOTTOM, HorizontalAlignment = Element.ALIGN_RIGHT, Colspan = 4, Rowspan = 2, Border = 0 });
                table.AddCell(new PdfPCell(new Phrase("บริษัท จันวาณิชย์ ซีเคียวริตี้พริ้นท์ติ้ง จำกัด", forNameTH)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Colspan = 6, Border = 0 });
                //--------------------
                table.AddCell(new PdfPCell(new Phrase("CHAN WANICH SECURITY PRINTING COMPANY LIMITED", forCSP)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Colspan = 6, Border = 0 });
                //--------------------
                table.AddCell(new PdfPCell(new Phrase("")) { VerticalAlignment = Element.ALIGN_BOTTOM, HorizontalAlignment = Element.ALIGN_RIGHT, Colspan = 4, Rowspan = 2, Border = 0 });
                table.AddCell(new PdfPCell(new Phrase("เลขประจำตัวผู้เสียภาษีอากร 0105533079571", forHeadOffice)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Colspan = 6, Border = 0 });
                //--------------------
                table.AddCell(new PdfPCell(new Phrase("สาขาที่ออกใบกำกับภาษี      00000 : สำนักงานใหญ่", forHeadOffice)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Colspan = 6, Border = 0 });
                //--------------------
                table.AddCell("");
                table.AddCell(new PdfPCell(new Phrase("699 ถนนสีลม แขวงสีลม เขตบางรัก\nกรุงเทพมหานคร 10500\nโทร.(662) 635-3333 โทรสาร (662) 236-7176\n699 SILOM ROAD, SILOM, BANGRAK\nBANGKOK 10500, THAILAND\nTEL.(662) 635-3333 FAX.(662) 236-7176", forNote)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, NoWrap = true, Border = 0, PaddingBottom = 10f });
                table.AddCell("");
                table.AddCell(new PdfPCell(new Phrase("ใบเสร็จรับเงิน/ใบกำกับภาษี     \nReceipt/Tax Invoice     \n", forNameTH)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Colspan = 4, Border = 0 });
                Phrase phrase = new Phrase();
                phrase.Add(new Chunk($"เลขที่/NO.  ", forHeadOffice));
                phrase.Add(new Chunk($"T03-0000000002", forDetails));
                phrase.Add(new Chunk($"\n"));
                phrase.Add(new Chunk($"วันที่/Date  ", forHeadOffice));
                phrase.Add(new Chunk($"16/07/2020", forDetails));
                table.AddCell(new PdfPCell(phrase)
                {
                    VerticalAlignment = Element.ALIGN_BOTTOM,
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    NoWrap = true,
                    Border = 0,
                    PaddingBottom = 10f
                });

                table.AddCell("");
                table.AddCell("");
                //--------------------
                table.AddCell(new PdfPCell(new Phrase(""))
                {
                    Border = 0,
                    PaddingBottom = 8f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                phrase = new Phrase();
                phrase.Add(new Chunk($"รหัสลูกค้า/Account Code ", forHeadOffice));
                phrase.Add(new Chunk($"004346", forDetails));
                table.AddCell(new PdfPCell(phrase)
                {
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    NoWrap = true,
                    Colspan = 3,
                    Border = 0,
                    BorderColorTop = new BaseColor(System.Drawing.Color.Black),
                    BorderWidthTop = 0.1f
                });

                phrase = new Phrase();
                phrase.Add(new Chunk($"TAX ID : ", forHeadOffice));
                phrase.Add(new Chunk($"0107557000489 : สำนักงานใหญ่", forDetails));
                table.AddCell(new PdfPCell(phrase)
                {
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    NoWrap = true,
                    Colspan = 3,
                    Border = 0,
                    BorderColorTop = new BaseColor(System.Drawing.Color.Black),
                    BorderWidthTop = 0.1f
                });
                phrase = new Phrase();
                phrase.Add(new Chunk($"ใบสั่งซื้อเลขที่/PO No. ", forHeadOffice));
                phrase.Add(new Chunk($"PO-00000001", forDetails));
                table.AddCell(new PdfPCell(phrase)
                {
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    Colspan = 3,
                    Border = 0,
                    BorderColorTop = new BaseColor(System.Drawing.Color.Black),
                    BorderWidthTop = 0.1f
                });
                //--------------------
                table.AddCell("");
                table.AddCell(new PdfPCell(new Phrase("ชื่อลูกค้า/Name", forHeaders)) { HorizontalAlignment = Element.ALIGN_LEFT, NoWrap = true, Border = 0 });
                table.AddCell(new PdfPCell(new Phrase("บริษัท ไอร่า แอนด์ ไอฟุล จำกัด (มหาชน)", forDetails)) { HorizontalAlignment = Element.ALIGN_LEFT, Colspan = 8, Border = 0 });
                //--------------------
                table.AddCell("");
                table.AddCell(new PdfPCell(new Phrase("ที่อยู่/Address", forHeaders)) { HorizontalAlignment = Element.ALIGN_LEFT, NoWrap = true, Border = 0 });
                table.AddCell(new PdfPCell(new Phrase("90 อาคาร ซีดับเบิ้ลยู ทาวเวอร์ เลขที่ห้องบี 3301-2 ชั้น33, 34ถนนรัชดภิเษกแขวงห้วยขวง เขตห้วยขวงกรุงเทพมหนคร 10310", forDetails)) { HorizontalAlignment = Element.ALIGN_LEFT, Colspan = 8, Border = 0, PaddingRight = 25f });
                //--------------------
                table.AddCell("");
                table.AddCell(new PdfPCell(new Phrase("สถานที่ส่งของ/Delivery To", forHeaders)) { HorizontalAlignment = Element.ALIGN_LEFT, NoWrap = true, Border = 0, Colspan = 2 });
                table.AddCell(new PdfPCell(new Phrase("90 อาคาร ซีดับเบิ้ลยู ทาวเวอร์ เลขที่ห้องบี 3301-2 ชั้น33, 34ถนนรัชดภิเษกแขวงห้วยขวง เขตห้วยขวงกรุงเทพมหนคร 10310", forDetails)) { HorizontalAlignment = Element.ALIGN_LEFT, Colspan = 7, Border = 0, PaddingRight = 25f });
                //--------------------
                table.AddCell(new PdfPCell(new Phrase("")) { Border = 0 });
                table.AddCell(new PdfPCell(new Phrase("วันครบกำหนด/Due Date", forHeaders)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, NoWrap = true, Colspan = 2, Border = 0 });
                table.AddCell(new PdfPCell(new Phrase("เงื่อนไขการชำระเงิน/Term of Payment", forHeaders)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, NoWrap = true, Colspan = 2, Border = 0 });
                table.AddCell(new PdfPCell(new Phrase("รหัสผู้ขาย/Salesman No.", forHeaders)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, NoWrap = true, Colspan = 2, Border = 0 });
                table.AddCell(new PdfPCell(new Phrase("ใบสั่งสินค้าเลขที่/Order No.", forHeaders)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, NoWrap = true, Colspan = 2, Border = 0 });
                table.AddCell(new PdfPCell(new Phrase("")) { Border = 0 });
                //--------------------

                table.AddCell(new PdfPCell(new Phrase("")) { Border = 0 });
                table.AddCell(new PdfPCell(new Phrase("15/08/2020", forDetails)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, NoWrap = true, Colspan = 2, Border = 0, PaddingBottom = 5f });
                table.AddCell(new PdfPCell(new Phrase("30 วันหลังจากส่งของ", forDetails)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, NoWrap = true, Colspan = 2, Border = 0, PaddingBottom = 5f });
                table.AddCell(new PdfPCell(new Phrase("ริญญารัตน์ เมธีกังวาลภัสร์", forDetails)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, NoWrap = true, Colspan = 2, Border = 0, PaddingBottom = 5f });
                table.AddCell(new PdfPCell(new Phrase("SCSP20-005593", forDetails)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, NoWrap = true, Colspan = 2, Border = 0, PaddingBottom = 5f });
                table.AddCell(new PdfPCell(new Phrase("")) { Border = 0 });
                //--------------------
                table.AddCell(new PdfPCell(new Phrase(""))
                {
                    Border = 0,
                    PaddingBottom = 8f,
                    BorderColorBottom = BaseColor.BLACK,
                    BorderWidthBottom = 0.1f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                table.AddCell(new PdfPCell(new Phrase("รายการสินค้าและบริการ\nDescription", forHeaders))
                {
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Colspan = 5,
                    Border = 0,
                    PaddingBottom = 8f,
                    BorderColorBottom = BaseColor.BLACK,
                    BorderWidthBottom = 0.1f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                table.AddCell(new PdfPCell(new Phrase("จำนวน\nQuantity", forHeaders))
                {
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Border = 0,
                    NoWrap = true,
                    PaddingBottom = 8f,
                    BorderColorBottom = BaseColor.BLACK,
                    BorderWidthBottom = 0.1f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                table.AddCell(new PdfPCell(new Phrase("หน่วยละ\nUnit Price", forHeaders))
                {
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Border = 0,
                    NoWrap = true,
                    PaddingBottom = 8f,
                    BorderColorBottom = BaseColor.BLACK,
                    BorderWidthBottom = 0.1f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                table.AddCell(new PdfPCell(new Phrase("จำนวนเงิน\nAmount", forHeaders))
                {
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Border = 0,
                    NoWrap = true,
                    PaddingBottom = 8f,
                    BorderColorBottom = BaseColor.BLACK,
                    BorderWidthBottom = 0.1f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                table.AddCell(new PdfPCell(new Phrase(""))
                {
                    Border = 0,
                    PaddingBottom = 8f,
                    BorderColorBottom = BaseColor.BLACK,
                    BorderWidthBottom = 0.1f,
                    BorderColorTop = BaseColor.BLACK,
                    BorderWidthTop = 0.1f
                });
                //----------------------
                table.AddCell(new PdfPCell(new Phrase("")) { Colspan = 10 });
                table.WriteSelectedRows(0, -1, document.LeftMargin, document.PageSize.Height - 30, writer.DirectContent);


                cb.BeginText();
                cb.SetFontAndSize(EnCodefont, 11);
                cb.SetTextMatrix(550 - len, document.PageSize.Height - 60);
                cb.ShowText(strPage);


                cb.EndText();
                cb.AddTemplate(template, 550, document.PageSize.Height - 60);

            }
            public override void OnCloseDocument(PdfWriter writer, Document document)
            {
                BaseFont EnCodefont = BaseFont.CreateFont(fontpath_normal, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                base.OnCloseDocument(writer, document);

                template.BeginText();
                template.SetFontAndSize(EnCodefont, 11);
                template.SetTextMatrix(0, 0);
                template.ShowText("" + (writer.PageNumber.ToString("N0")));
                template.EndText();

            }
        }
        #endregion

        #region TrueMove_Credit_Note
        public FileResponseDataBinding CreateTrueMoveCDN(Refund creditNoteModel)
        {
            var response = new FileResponseDataBinding();
            PdfPTable table = null;
            PdfWriter writer = null;
            PdfPCell cell = null;
            PdfPTable tbDetails = null;
            PdfPTable tbFooter = null;

            try
            {
                Document pdfDoc = new Document(PageSize.A4, 40, 40, 30, 50);

                using (MemoryStream ms = new MemoryStream())
                {
                    writer = PdfWriter.GetInstance(pdfDoc, ms);
                    string fontpath_normal = Path.Combine(Environment.CurrentDirectory, "Fonts", "THSarabunNew.ttf");
                    string fontpathBold = Path.Combine(Environment.CurrentDirectory, "Fonts", "THSarabunNew Bold.ttf");
                    BaseFont EnCodefont = BaseFont.CreateFont(fontpath_normal, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    BaseFont EnCodefontBold = BaseFont.CreateFont(fontpathBold, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    Font formalFont = new Font(EnCodefont, 16, Font.NORMAL);
                    Font formalBold = new Font(EnCodefont, 16, Font.BOLD);
                    Font forCreditNote = new Font(EnCodefontBold, 19, Font.NORMAL);
                    Font forCATreseller = new Font(EnCodefont, 17, Font.BOLD, BaseColor.LIGHT_GRAY);
                    Font forHeadOffice = new Font(EnCodefontBold, 12, Font.NORMAL, BaseColor.BLACK);
                    Font forAddrCustomer = new Font(EnCodefont, 12, Font.NORMAL, BaseColor.DARK_GRAY);
                    Font forHeadDetails = new Font(EnCodefontBold, 11, Font.NORMAL, BaseColor.BLACK);
                    Font forDetails = new Font(EnCodefont, 10, Font.NORMAL, BaseColor.DARK_GRAY);
                    Font forReason = new Font(EnCodefont, 13, Font.NORMAL, BaseColor.DARK_GRAY);
                    Font forNameCustomer = new Font(EnCodefont, 15, Font.NORMAL, BaseColor.DARK_GRAY);

                    ////---- gen pdf to currentPath
                    //PdfWriter.GetInstance(pdfDoc, new FileStream(Environment.CurrentDirectory + "/apiGen.pdf", FileMode.Create, FileAccess.Write));

                    pdfDoc.Open();

                    Image image = Image.GetInstance(Environment.CurrentDirectory + "/Images/Truemove.jpg");
                    image.ScaleAbsolute(150, 25);
                    image.Alignment = Image.ALIGN_LEFT;
                    pdfDoc.Add(image);

                    table = new PdfPTable(7);
                    table.TotalWidth = 500f;
                    table.LockedWidth = true;
                    table.SetWidths(new float[] { 125f, 10f, 110f, 110f, 35f, 10f, 100f });
                    //table.DefaultCell.FixedHeight = 200f;
                    table.HeaderRows = 10;

                    table.AddCell(new PdfPCell(new Phrase(" Credit note", forCreditNote)) { VerticalAlignment = Element.ALIGN_CENTER, HorizontalAlignment = Element.ALIGN_LEFT, NoWrap = true, Border = 0, PaddingTop = 10 });
                    table.AddCell(new PdfPCell(new Phrase("", forCreditNote)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0, PaddingTop = 10 });
                    table.AddCell(new PdfPCell(new Phrase("", forCreditNote)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0, PaddingTop = 10 });
                    table.AddCell(new PdfPCell(new Phrase("", forCreditNote)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0, PaddingTop = 10 });
                    table.AddCell(new PdfPCell(new Phrase("", forCreditNote)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0, PaddingTop = 10 });
                    table.AddCell(new PdfPCell(new Phrase("A CAT reseller", forCATreseller))
                    {
                        VerticalAlignment = Element.ALIGN_CENTER,
                        HorizontalAlignment = Element.ALIGN_RIGHT,
                        Colspan = 2,
                        PaddingTop = 10,
                        NoWrap = true,
                        Border = 0
                    });

                    table.AddCell(new PdfPCell(new Phrase("", forCreditNote)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", formalBold)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", formalBold)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", formalBold)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", formalBold)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase($"Head Office\n" +
                        $"{creditNoteModel.Seller?.SellerName}\n" +
                        $"{creditNoteModel.Seller?.BuildingNumber} {creditNoteModel.Seller?.BuildingName} {(string.IsNullOrWhiteSpace(creditNoteModel.Seller?.StreetName) ? "" : creditNoteModel.Seller?.StreetName + " Road")}\n" +
                        $"{creditNoteModel.Seller?.CitySubDivisionName} {creditNoteModel.Seller?.CityName} {creditNoteModel.Seller?.CountrySubDivisionName} {creditNoteModel.Seller?.PostCode}", forHeadOffice))
                    {
                        VerticalAlignment = Element.ALIGN_MIDDLE,
                        HorizontalAlignment = Element.ALIGN_RIGHT,
                        Colspan = 2,
                        PaddingBottom = 10,
                        NoWrap = true,
                        Border = 0
                    });

                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("Date", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase(":", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase($"{creditNoteModel.ActivityDate}", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });

                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("Number", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase(":", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase($"{creditNoteModel.DocumentNumber}", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });

                    table.AddCell(new PdfPCell(new Phrase("Account Number", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase(":", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase($"{creditNoteModel.AccountNumber}", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });

                    table.AddCell(new PdfPCell(new Phrase("Service No./Number of Product", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase(":", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });

                    table.AddCell(new PdfPCell(new Phrase("Customer Tax ID", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase(":", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase($"{creditNoteModel.Buyer?.BuyerTaxId}", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });

                    table.AddCell(new PdfPCell(new Phrase("Branch No.", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase(":", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase($"{creditNoteModel.Buyer?.BranchNo}", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });

                    table.AddCell(new PdfPCell(new Phrase("VAT Address", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase(":", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase($"{creditNoteModel.Buyer?.BuyerName}", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });

                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase("", forAddrCustomer)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 });
                    table.AddCell(new PdfPCell(new Phrase(
                        $"{creditNoteModel.Buyer?.BuildingNumber}" +
                        $"{(string.IsNullOrWhiteSpace(creditNoteModel.Buyer?.Moo) ? " " : " Moo " + creditNoteModel.Buyer?.Moo)} " +
                        $"Soi {creditNoteModel.Buyer?.Soi} " +
                        $"Subsoi {creditNoteModel.Buyer?.SubSoi} " +
                        $"Village/Building {creditNoteModel.Buyer?.BuildingName} " +
                        $"Floor {creditNoteModel.Buyer?.Floor} " +
                        $"Room {creditNoteModel.Buyer?.RoomNo} {creditNoteModel.Buyer?.StreetName} " +
                        $"Road {creditNoteModel.Buyer?.CitySubDivisionName} {creditNoteModel.Buyer?.CityName} {creditNoteModel.Buyer?.CountrySubDivisionName} {creditNoteModel.Buyer?.PostCode}", forAddrCustomer))
                    {
                        VerticalAlignment = Element.ALIGN_MIDDLE,
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        Colspan = 5,
                        Border = 0,
                        PaddingBottom = 20
                    });

                    cell = new PdfPCell(new Phrase("")) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Border = 0 };
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("")) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Border = 0 };
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("")) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Border = 0 };
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("")) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Border = 0 };
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("")) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Border = 0 };
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("")) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Border = 0 };
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("")) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER, Border = 0 };
                    table.AddCell(cell);

                    pdfDoc.Add(table);

                    tbDetails = new PdfPTable(4);
                    tbDetails.TotalWidth = 500f;
                    tbDetails.LockedWidth = true;
                    tbDetails.SetWidths(new float[] { 175f, 125f, 105f, 100f });
                    //tbDetails.HeaderRows = 2;
                    cell = new PdfPCell(new Phrase("Description", forHeadDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0, BackgroundColor = BaseColor.GRAY, PaddingBottom = 6 };
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Amount (Baht)", forHeadDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0, BackgroundColor = BaseColor.GRAY, PaddingBottom = 6 };
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase("VAT (Baht)", forHeadDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0, BackgroundColor = BaseColor.GRAY, PaddingBottom = 6 };
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Total (Baht)", forHeadDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0, BackgroundColor = BaseColor.GRAY, PaddingBottom = 6 };
                    tbDetails.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    cell.Colspan = 4;
                    tbDetails.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Original Receipt Number", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    cell.PaddingLeft = 5;
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase($"{creditNoteModel.RNAmount?.Amount}", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 };
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase($"{creditNoteModel.RNAmount?.VatAmount}", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 };
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase($"{creditNoteModel.RNAmount?.TotalAmount}", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 };
                    tbDetails.AddCell(cell);

                    cell = new PdfPCell(new Phrase($"{creditNoteModel.OriginalReceiptNumber}", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    cell.PaddingLeft = 5;
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase(" ", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 };
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase(" ", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 };
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase(" ", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 };
                    tbDetails.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Corrected Amount", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    cell.PaddingLeft = 5;
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase($"{creditNoteModel.CorrAmount?.Amount}", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 };
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase($"{creditNoteModel.CorrAmount?.VatAmount}", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 };
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase($"{creditNoteModel.CorrAmount?.TotalAmount}", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 };
                    tbDetails.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Different Amount", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    cell.PaddingLeft = 5;
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase($"{creditNoteModel.DiffAmount?.Amount}", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 };
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase($"{creditNoteModel.DiffAmount?.VatAmount}", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 };
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase($"{creditNoteModel.DiffAmount?.TotalAmount}", forDetails)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_RIGHT, Border = 0 };
                    cell.PaddingBottom = 20;
                    tbDetails.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Reason for Adjustment : No service performed according to the Contract", forReason)) { VerticalAlignment = Element.ALIGN_CENTER, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    cell.Colspan = 4;
                    tbDetails.AddCell(cell);

                    cell = new PdfPCell(new Phrase("This document has been prepared and submitted to the Revenue Department by electronic means. It can be used for legal purpose.", forDetails)) { VerticalAlignment = Element.ALIGN_CENTER, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    cell.PaddingTop = 50;
                    cell.Colspan = 4;
                    tbDetails.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Please check the accuracy of the detail within 7 days of receiving, otherwise it shall be deemed that this document is accurate and complete.", forDetails)) { VerticalAlignment = Element.ALIGN_CENTER, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    cell.PaddingBottom = 50;
                    cell.Colspan = 4;
                    tbDetails.AddCell(cell);

                    pdfDoc.Add(tbDetails);

                    pdfDoc.Add(image);

                    tbFooter = new PdfPTable(1);
                    tbFooter.TotalWidth = 500f;
                    tbFooter.LockedWidth = true;
                    cell = new PdfPCell(new Phrase($"{creditNoteModel.Seller?.SellerName}\n" +
                        $"{creditNoteModel.Seller?.BuildingNumber} {creditNoteModel.Seller?.BuildingName} {(string.IsNullOrWhiteSpace(creditNoteModel.Seller?.StreetName) ? "" : creditNoteModel.Seller?.StreetName + " Road")}\n" +
                        $"{creditNoteModel.Seller?.CitySubDivisionName} {creditNoteModel.Seller?.CityName} {creditNoteModel.Seller?.CountrySubDivisionName} {creditNoteModel.Seller?.PostCode}", forHeadOffice))
                    { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    cell.PaddingTop = 15;
                    cell.PaddingBottom = 15;
                    tbFooter.AddCell(cell);

                    cell = new PdfPCell(new Phrase($"{creditNoteModel.Buyer?.BuyerName}", forNameCustomer)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    tbFooter.AddCell(cell);
                    cell = new PdfPCell(new Phrase($"{creditNoteModel.Buyer?.BuildingNumber}" +
                        $"{(string.IsNullOrWhiteSpace(creditNoteModel.Buyer?.Moo) ? " " : " Moo " + creditNoteModel.Buyer?.Moo)} " +
                        $"Soi {creditNoteModel.Buyer?.Soi} " +
                        $"Subsoi {creditNoteModel.Buyer?.SubSoi} ", forNameCustomer))
                    { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    tbFooter.AddCell(cell);
                    cell = new PdfPCell(new Phrase(
                        $"Village/Building {creditNoteModel.Buyer?.BuildingName} " +
                        $"Floor {creditNoteModel.Buyer?.Floor} " +
                        $"Room {creditNoteModel.Buyer?.RoomNo} {creditNoteModel.Buyer?.StreetName} Road", forNameCustomer))
                    { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    tbFooter.AddCell(cell);
                    cell = new PdfPCell(new Phrase($"{creditNoteModel.Buyer?.CitySubDivisionName} {creditNoteModel.Buyer?.CityName}", forNameCustomer)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    tbFooter.AddCell(cell);
                    cell = new PdfPCell(new Phrase($"{creditNoteModel.Buyer?.CountrySubDivisionName} {creditNoteModel.Buyer?.PostCode}", forNameCustomer)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    tbFooter.AddCell(cell);
                    cell = new PdfPCell(new Phrase("A CAT reseller", forCATreseller)) { VerticalAlignment = Element.ALIGN_TOP, HorizontalAlignment = Element.ALIGN_LEFT, Border = 0 };
                    tbFooter.AddCell(cell);

                    pdfDoc.Add(tbFooter);

                    pdfDoc.Close();
                    response.File = ms.ToArray();
                    response.StatusCode = 200;
                }
            }
            catch (Exception ex)
            {
                response.StatusCode = 500;
                response.Message = ex.Message;
                return response;
            }

            return response;
        }


        #endregion


        public static readonly string fileDll = "EBTS.Security.Cryptography.Library.dll";
        public static readonly string pathLocation = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
        public ResponseModel EncrptPassword(string plainText)
        {
            try
            {
                if (string.IsNullOrEmpty(plainText)) return new ResponseModel() { Result = "string is null or empty." };

                var dllPath = Path.Combine(pathLocation, fileDll);

                using (FileStream fs = System.IO.File.OpenRead(Path.Combine(dllPath)))
                {
                    using (HashAlgorithm hashAlgorithm = SHA256.Create())
                    {
                        byte[] hash = hashAlgorithm.ComputeHash(fs);

                        var plainTextBytes = Encoding.UTF8.GetBytes(plainText);

                        using (Aes encryptor = Aes.Create())
                        {
                            Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(ToHexString(hash), hash);
                            encryptor.Key = pdb.GetBytes(32);
                            encryptor.IV = pdb.GetBytes(16);
                            using (MemoryStream ms = new MemoryStream())
                            {
                                using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                                {
                                    cs.Write(plainTextBytes, 0, plainTextBytes.Length);
                                    cs.Close();
                                }
                                plainText = Convert.ToBase64String(ms.ToArray());
                                return new ResponseModel() { Result = plainText };
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return new ResponseModel() { Result = ex.Message };
            }
        }

        public ResponseModel DecrptPassword(string cipherText)
        {
            try
            {
                if (string.IsNullOrEmpty(cipherText)) return new ResponseModel() { Result = "string is null or empty." };

                var dllPath = Path.Combine(pathLocation, fileDll);

                using (FileStream fs = System.IO.File.OpenRead(Path.Combine(dllPath)))
                {
                    using (HashAlgorithm hashAlgorithm = SHA256.Create())
                    {
                        byte[] hash = hashAlgorithm.ComputeHash(fs);

                        byte[] cipherBytes = Convert.FromBase64String(cipherText);

                        using (Aes encryptor = Aes.Create())
                        {
                            Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(ToHexString(hash), hash);
                            encryptor.Key = pdb.GetBytes(32);
                            encryptor.IV = pdb.GetBytes(16);
                            using (MemoryStream ms = new MemoryStream())
                            {
                                using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                                {
                                    cs.Write(cipherBytes, 0, cipherBytes.Length);
                                    cs.Close();
                                }
                                cipherText = Encoding.UTF8.GetString(ms.ToArray());
                                return new ResponseModel() { Result = cipherText };
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return new ResponseModel() { Result = ex.Message };
            }
        }

        public static string ToHexString(byte[] bytes)
        {
            StringBuilder sb = new StringBuilder();
            foreach (byte b in bytes)
            {
                sb.Append(b.ToString("x2").ToLower());
            }
            return sb.ToString();
        }
    }
}
