using DocXToPDF;

namespace TestProject1
{
    public class UnitTest
    {
        [Fact]
        public void CreatePdfTest()
        {
            Converter.LibreOfficePath = @"C:\Program Files\LibreOffice\program\soffice.exe";
            Converter.TimeOut = 60000;
            string docxWay = @"FileToBeConverted.docx";
            string pdfWay = @"D:\ConvertedFile.pdf";
            byte[] docxBytes = File.ReadAllBytes(docxWay);
            using (MemoryStream ms = new MemoryStream(docxBytes))
            {
                byte[]? pdfBytes = Converter.ConvertToPdf(ms);
                if (pdfBytes != null)
                    File.WriteAllBytes(pdfWay, pdfBytes);
            }
        }
    }
}