# Библиотека Converter: Конвертация DOCX в PDF

Простая библиотека для преобразования документов DOCX в формат PDF с использованием LibreOffice.

## Предварительные требования

1. **LibreOffice**  
   Установите последнюю версию LibreOffice:  
   [https://www.libreoffice.org/download](https://www.libreoffice.org/download)

2. **.NET 8.0**  
   Убедитесь, что установлена среда выполнения .NET 8.0:  
   [https://dotnet.microsoft.com/download](https://dotnet.microsoft.com/download)

## Настройка

Перед использованием задайте параметры в коде приложения:

```csharp
// Укажите путь к исполняемому файлу LibreOffice
Converter.LibreOfficePath = @"C:\Program Files\LibreOffice\program\soffice.exe";

// Опционально: установите таймаут в миллисекундах (по умолчанию 30000 мс)
Converter.TimeOut = 60000;
```
Пример использования
```csharp
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

```
