using System;
using System.Windows.Forms;

using System.IO;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Barcodes;
using iText.Kernel.Pdf.Xobject;
using iText.IO.Image;
using iText.Kernel.Geom;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.IO.Font;
using iText.Layout.Borders;

namespace MyPdf
{

    public partial class Form1 : Form
    {
        public static readonly string DEST = "C:/Neosmart Solutions Inc/PDFs/result.pdf";
        int paper_width = 336, paper_height = 1056;
        int img_size_width = 100, img_size_height = 100;
        int offset_x = 2;

        public Form1()
        {
            InitializeComponent();
        }

        private void btn_create_Click(object sender, EventArgs e)
        {
            FileInfo file = new FileInfo(DEST);
            file.Directory.Create();

            new Form1().ManipulatePdf(DEST);
        }
        private void ManipulatePdf(string dest)
        {
            PdfDocument pdfDoc = new PdfDocument(new PdfWriter(dest));
            Document doc = new Document(pdfDoc);
            doc.SetMargins(0f, 0f, 0f, 0f);

            String img_path = "../../resource/image/1.png";
            String FONT = "../../resource/font/arial.ttf";

            Image image_1 = new Image(ImageDataFactory.Create(img_path));
            Image image_2 = new Image(ImageDataFactory.Create(img_path));
            image_1.ScaleAbsolute(img_size_width, img_size_height);
            image_2.ScaleAbsolute(img_size_width, img_size_height);

            pdfDoc.AddNewPage(new PageSize(paper_width, paper_height));

            image_1.SetFixedPosition(1, offset_x, paper_height - img_size_height);
            doc.Add(image_1);
            image_2.SetFixedPosition(1, offset_x + paper_width - img_size_width, paper_height - img_size_height);
            doc.Add(image_2);

            PdfFont TitleFont = PdfFontFactory.CreateFont(FONT, null, true);

            Paragraph p = new Paragraph().SetTextAlignment(TextAlignment.CENTER);
            Text text_header_1 = new Text("MONTEREY");
            //PdfFont TitleFont = PdfFontFactory.CreateFont(FontConstants.COURIER);
            text_header_1.SetFont(TitleFont).SetBold().SetItalic().SetFontSize(20);
            p.Add(text_header_1);
            doc.Add(p);

            p = new Paragraph().SetTextAlignment(TextAlignment.CENTER);
            Text text_header_2 = new Text("BOTTLE DEPOT");
            text_header_2.SetFont(TitleFont).SetBold().SetItalic().SetFontSize(13);
            p.SetMultipliedLeading(-0.8f);
            p.Add(text_header_2);
            doc.Add(p);

            p = new Paragraph().SetTextAlignment(TextAlignment.CENTER);
            Text text_header_3 = new Text("2240 - 68 ST NE");
            text_header_3.SetFont(TitleFont).SetBold().SetFontSize(9);
            p.SetMultipliedLeading(0);
            p.Add(text_header_3);
            doc.Add(p);

            p = new Paragraph().SetTextAlignment(TextAlignment.CENTER);
            text_header_3 = new Text("CALGARY, AB T1Y 0A2");
            text_header_3.SetFont(TitleFont).SetBold().SetFontSize(9);
            p.SetMultipliedLeading(0);
            p.Add(text_header_3);
            doc.Add(p);

            p = new Paragraph().SetTextAlignment(TextAlignment.CENTER);
            text_header_3 = new Text("PH:(403) 590-6665");
            text_header_3.SetFont(TitleFont).SetBold().SetFontSize(9);
            p.SetMultipliedLeading(0);
            p.Add(text_header_3);
            doc.Add(p);

            PdfFont special_Font = PdfFontFactory.CreateFont(FONT, PdfEncodings.IDENTITY_H, true);
            p = new Paragraph().SetTextAlignment(TextAlignment.CENTER);
            Text text_header_4 = new Text("* * P A I D - R E P R I N T * *");
            text_header_4.SetFont(special_Font).SetBold().SetFontSize(14);
            p.SetMultipliedLeading(2f);
            p.Add(text_header_4);
            doc.Add(p);


            String code = "741235";
            Barcode128 code128 = new Barcode128(pdfDoc);
            code128.SetFont(null);
            code128.SetCode(code);
            code128.SetCodeType(Barcode128.CODE128);
            Image code128Image = new Image(code128.CreateFormXObject(pdfDoc)).ScaleAbsolute(paper_width * 0.4f, paper_width * 0.15f);
            Paragraph paragraph = new Paragraph(code).SetTextAlignment(TextAlignment.CENTER);
            code128Image.SetMarginLeft(paper_width * 0.29f);
            Cell cell = new Cell();
            cell.Add(code128Image);
            cell.Add(paragraph);
            doc.Add(cell);

            string[] org_lines = System.IO.File.ReadAllLines("../../resource/input_table.csv");
            Table table = new Table(UnitValue.CreatePercentArray(new float[] { 3, 3, 3 })).UseAllAvailableWidth();
            table.SetTextAlignment(TextAlignment.CENTER);
            table.SetMargin(15f);
            for (int i = 0; i < org_lines.Length-1; i++)
            {
                string[] temp = org_lines[i].Split(',');
                for(int j = 0; j < 3; j++)
                {
                    table.AddCell(temp[j]);
                }
            }
            doc.Add(table);

            p = new Paragraph().SetTextAlignment(TextAlignment.CENTER);
            string[] total_temp = org_lines[org_lines.Length-1].Split(',');
            text_header_3 = new Text(total_temp[1] + "  " + total_temp[2]);
            text_header_3.SetFont(TitleFont).SetBold().SetFontSize(15);
            p.SetMultipliedLeading(0);
            p.Add(text_header_3);
            doc.Add(p);


            p = new Paragraph().SetTextAlignment(TextAlignment.CENTER);
            DateTime localDate = DateTime.Now;
            string format = "MMM d, yyyy\t\t\t\t\t";
            text_header_3 = new Text(localDate.ToString(format) + localDate.ToString("t"));
            text_header_3.SetFont(TitleFont).SetFontSize(12);
            p.SetMultipliedLeading(2);
            p.Add(text_header_3);
            doc.Add(p);

            p = new Paragraph().SetTextAlignment(TextAlignment.CENTER);
            text_header_3 = new Text("You were served at MONTEREY-6 by Subash.");
            text_header_3.SetFont(TitleFont).SetFontSize(12);
            p.SetMultipliedLeading(2);
            p.Add(text_header_3);
            doc.Add(p);

            p = new Paragraph().SetTextAlignment(TextAlignment.CENTER);
            text_header_3 = new Text("Thank you for your patronage.");
            text_header_3.SetFont(TitleFont).SetBold().SetFontSize(12);
            p.SetMultipliedLeading(1);
            p.Add(text_header_3);
            doc.Add(p);

            doc.Close();
            MessageBox.Show("File saved to C:\\Neosmart Solutions Inc\\PDFs\\result.pdf correctly!");
        }

    }
}
