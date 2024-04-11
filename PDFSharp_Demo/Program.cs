using System.Diagnostics;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using MigraDocCore.DocumentObjectModel;
using PdfSharpCore.Pdf.IO;

#region 測試一，使用 PdfSharpCore 產生 PDF (畫布繪圖)
// REF: https://github.com/empira/PDFsharp.Samples/tree/master
var cover = new PdfDocument();

cover.Info.Title = "Created with PDFsharp";
cover.Info.Subject = "Just a simple Hello-World program.";

var page = cover.AddPage();
// 用繪圖方式製作 PDF (另有 MigraDoc 提供文件模型)
var gfx = XGraphics.FromPdfPage(page);
var width = page.Width;
var height = page.Height;

// 背景填色
gfx.DrawRectangle(XBrushes.SteelBlue, 0, 0, width, height);

// 繪製矩形，灰框白底
double margin = 80;
XPen grayPen = new(XColors.Gray, 4);

gfx.DrawRectangle(grayPen, XBrushes.White, margin, margin, width - 2 * margin, 150);

// 繪製文字
var font = new XFont("Tahoma", 32, XFontStyle.Bold);

gfx.DrawString("THIS IS A BOOK.", font, XBrushes.DarkGray, new XRect(margin, margin, width - 2 * margin, 150), XStringFormats.Center);

var coverFilePath = Path.GetTempFileName() + ".pdf";

cover.Save(coverFilePath);

// 開啟檔案
Process.Start(new ProcessStartInfo(coverFilePath) { UseShellExecute = true });
#endregion

#region 測試二，使用 MigraDocCore 產生 PDF (文件模型)
//REF: https://github.com/empira/PDFsharp.Samples/blob/master/src/samples/src/MigraDoc/src/HelloMigraDoc/Styles.cs
var document = new Document()
{
    Info = {
        Title = "Created with MigraDoc",
        Subject = "Hello, World!",
        Author = "Jeffrey Lee"
    }
};
var style = document.Styles["Normal"] ?? throw new InvalidOperationException("Style Normal not found.");

style.Font.Name = "Segoe UI";
style = document.Styles["Heading1"];
style.Font.Size = 16;
style.Font.Bold = true;
style.Font.Color = Colors.DarkBlue;
style.ParagraphFormat.PageBreakBefore = true;
style.ParagraphFormat.SpaceAfter = 6;
style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
style.ParagraphFormat.KeepWithNext = true;

var sec = document.AddSection();
var para = sec.AddParagraph();

para.AddFormattedText("Header Test", "Heading1");
para = sec.AddParagraph();
para.AddLineBreak();
para.AddText("Hello, World!");

var pdfRenderer = new MigraDocCore.Rendering.PdfDocumentRenderer()
{
    Document = document,
    PdfDocument = new PdfDocument()
};

pdfRenderer.RenderDocument();

var pdfFilePath = Path.GetTempFileName() + ".pdf";

pdfRenderer.PdfDocument.Save(pdfFilePath);

// 開啟檔案
Process.Start(new ProcessStartInfo(pdfFilePath) { UseShellExecute = true });
#endregion

#region 測試三，合併兩份 PDF
var pdf1 = PdfReader.Open(coverFilePath, PdfDocumentOpenMode.Import);
var pdf2 = PdfReader.Open(pdfFilePath, PdfDocumentOpenMode.Import);
var pdf3 = new PdfDocument();

for (int i = 0; i < pdf1.PageCount; i++)
{
    pdf3.AddPage(pdf1.Pages[i]);
}

for (int i = 0; i < pdf2.PageCount; i++)
{
    pdf3.AddPage(pdf2.Pages[i]);
}

var mergedFilePath = Path.GetTempFileName() + ".pdf";

pdf3.Save(mergedFilePath);

// 開啟檔案
Process.Start(new ProcessStartInfo(mergedFilePath) { UseShellExecute = true });
#endregion

#region 測試四，為 PDF 加上浮水印
// https://www.pdfsharp.net/wiki/Watermark-sample.ashx
var mergedPdf = PdfReader.Open(mergedFilePath, PdfDocumentOpenMode.Modify);
var wartermarkText = "TOP SECRET";
var formatWM = new XStringFormat()
{
    Alignment = XStringAlignment.Near,
    LineAlignment = XLineAlignment.Near
};
var fontWM = new XFont("Tahoma", 48, XFontStyle.Bold);
XBrush brushWM = new XSolidBrush(XColor.FromArgb(96, 255, 0, 0));

foreach (var pg in mergedPdf.Pages)
{
    var gfxWM = XGraphics.FromPdfPage(pg, XGraphicsPdfPageOptions.Append);
    var size = gfxWM.MeasureString(wartermarkText, fontWM);

    gfxWM.TranslateTransform(pg.Width.Point / 2, pg.Height.Point / 2);
    gfxWM.RotateTransform(-Math.Atan(pg.Height / pg.Width) * 180 / Math.PI);
    gfxWM.TranslateTransform(-pg.Width.Point / 2, -pg.Height.Point / 2);
    gfxWM.DrawString(wartermarkText, fontWM, brushWM, new XPoint((pg.Width.Point - size.Width) / 2, (pg.Height.Point - size.Height) / 2), formatWM);
}

var watermarkedFilePath = Path.GetTempFileName() + ".pdf";

mergedPdf.Save(watermarkedFilePath);

// 開啟檔案
Process.Start(new ProcessStartInfo(watermarkedFilePath) { UseShellExecute = true });
#endregion