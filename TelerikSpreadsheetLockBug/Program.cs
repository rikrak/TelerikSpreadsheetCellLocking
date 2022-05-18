using System;
using System.IO;
using Telerik.Documents.Common.Model;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders.OpenXml.Xlsx;
using Telerik.Windows.Documents.Spreadsheet.Model;
using Telerik.Windows.Documents.Spreadsheet.Model.Protection;

namespace TelerikSpreadsheetLockBug
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (var wb = new Workbook())
            {
                var sheet = wb.Worksheets.Add();
                sheet.Name = "LockedFillDemo";
                
                var a1Cell = sheet.Cells[0, 0];
                a1Cell.SetValue("this is locked");

                // When opening the spreadsheet in Excel, cell B1 will appear with a bright yellow background
                // if you try and edit it, you will receive a warning that it is protected, even though it is explicitly set to unlocked
                var b1Cell = sheet.Cells[0, 1];
                var yellow = FromHtml("#FFFF00");
                var solidFill = PatternFill.CreateSolidFill(yellow);
                b1Cell.SetFill(solidFill);
                b1Cell.SetIsLocked(false);  // explicitly unlock the cell

                // this line seems to cause the issue
                // without it, cell B1 is editable when loaded in Excel
                sheet.Rows[0].SetIsLocked(false);  // unlock the whole row that B1 is on

                var password = "spreadsheet";
                var protectionOptions = BuildProtectionOptions();
                sheet.Protect(password, protectionOptions);

                var fileName = $"{DateTime.Now:yyyy-MM-dd-hhmmss}_sheet.xlsx";

                var xlsxFormatter = new XlsxFormatProvider();
                var b = xlsxFormatter.Export(wb);
                File.WriteAllBytes(fileName, b);
            }
        }

        private static WorksheetProtectionOptions BuildProtectionOptions()
        {
            var protectionOptions = new WorksheetProtectionOptions(
                allowDeleteRows: true,
                allowInsertRows: true,
                allowFormatCells: true,
                allowFormatRows: true,
                allowFiltering: true,
                allowSorting: true,
                allowDeleteColumns: false,
                allowInsertColumns: false,
                allowFormatColumns: true
            );
            return protectionOptions;
        }

        private static ThemableColor FromHtml(string htmlRgb)
        {
            if (string.IsNullOrWhiteSpace(htmlRgb)) throw new ArgumentException("Value cannot be null or whitespace.", nameof(htmlRgb));
            var rgb = htmlRgb.Replace('#', ' ').Trim();
            if (rgb.Length == 3)
            {
                rgb = $"{rgb[0]}{rgb[0]}{rgb[1]}{rgb[1]}{rgb[2]}{rgb[2]}";
            }

            if (rgb.Length != 6)
            {
                throw new ArgumentException($"The RGB colour {htmlRgb} is in an invalid format", nameof(htmlRgb));
            }

            var rgbBytes = StringToByteArray(rgb);

            return ThemableColor.FromArgb(Byte.MaxValue, rgbBytes[0], rgbBytes[1], rgbBytes[2]);
        }

        private static byte[] StringToByteArray(string hex)
        {
            int numberChars = hex.Length;
            byte[] bytes = new byte[numberChars / 2];
            for (int i = 0; i < numberChars; i += 2)
            {
                bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
            }
            return bytes;
        }

    }
}
