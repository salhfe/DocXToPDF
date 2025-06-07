using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocXToPDF
{
    public static class Converter
    {
        public static string LibreOfficePath = "";
        public static int TimeOut = 30000;
        public static byte[]? ConvertToPdf(MemoryStream docxStream)
        {
            if (LibreOfficePath == string.Empty) // если не указали путь к LibreOffice, возвращем null
                return null;
            string tempDir = Path.GetTempPath(); // возвращает путь к временной директории (на линуксе должно работать)
            string inputFilePath = Path.Combine(tempDir, $"{Guid.NewGuid()}.docx"); //  Создание пути для временного файла .docx
            string outputDir = Path.Combine(tempDir, Guid.NewGuid().ToString());
            // Guid.NewGuid() создаёт рандомный индентификатор

            Directory.CreateDirectory(outputDir);

            try
            {
                // Сохраняем MemoryStream во временный файл
                using (var fileStream = new FileStream(inputFilePath, FileMode.Create))
                {
                    docxStream.CopyTo(fileStream);
                }

                // Выполняем конвертацию
                Convert(inputFilePath, outputDir);

                // Ищем сконвертированный PDF
                string? pdfPath = Directory.GetFiles(outputDir, "*.pdf").FirstOrDefault();
                if (pdfPath == null)
                    throw new FileNotFoundException("Не найден PDF файл после конвертации");

                // Загружаем PDF в MemoryStream
                var pdfStream = new MemoryStream();
                using (var fileStream = new FileStream(pdfPath, FileMode.Open))
                {
                    fileStream.CopyTo(pdfStream);
                }
                pdfStream.Position = 0; // Сбрасываем позицию потока

                return pdfStream.ToArray();
            }
            finally
            {
                // Удаляем временные файлы
                if (File.Exists(inputFilePath)) File.Delete(inputFilePath);
                if (Directory.Exists(outputDir)) Directory.Delete(outputDir, true);
            }
        }

        private static void Convert(string inputPath, string outputDir)
        {
            var processInfo = new ProcessStartInfo
            {
                FileName = LibreOfficePath,
                Arguments = $"--headless --convert-to pdf --outdir \"{outputDir}\" \"{inputPath}\"", // headless - LibreOffice работает без интерфейса
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using (var process = new Process())
            {
                process.StartInfo = processInfo;
                process.Start();

                // Читаем вывод для диагностики
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();

                if (!process.WaitForExit(TimeOut))
                    throw new Exception("LibreOffice слишком долго конвертировал");

                if (process.ExitCode != 0)
                    throw new Exception($"Конвертация неудалась: {error}\nЛоги: {output}");
            }
        }
    }
}