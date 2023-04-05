using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace VooDooRPA_Project
{
    class Program
    {
        private static readonly string ekstrePath = @"C:\voodoorpa\EKSTRE-GIRDI.xlsx";
        private static readonly string formulPath = @"C:\voodoorpa\FORMUL-GIRDI.xlsx";
        private static readonly string raporPath = @"C:\voodoorpa\RAPOR.xlsx";

        private static readonly string islemSonucuWorksheetName = "İşlem Sonucu";
        private static readonly string pivotRaporWorksheetName = "Pivot Rapor";

        private static readonly string siraNoHeader = "SIRA_NO";
        private static readonly string adetHeader = "ADET";
        private static readonly string kgDesiHeader = "KG_DESI";
        private static readonly string mesafeHeader = "MESAFE";
        private static readonly string artanHeader = "Artan Her Desi İçin";
        private static readonly string ucretHeader = "ÜCRET";

        private static ExcelHandler excelHandler = new ExcelHandler();

        static void Main(string[] args)
        {
            ExcelPackage ekstre = excelHandler.GetExcelPackage(ekstrePath);
            FileCheck(ekstre, ekstrePath);
            ExcelPackage formul = excelHandler.GetExcelPackage(formulPath);
            FileCheck(formul, formulPath);
            Console.WriteLine();

            ExcelPackage rapor = CreateRapor(ekstre, formul);
            Console.WriteLine("Çıkmak için bir tuşa basınız. ");
            Console.ReadKey();
            if (File.Exists(raporPath))
            {
                Process.Start(new ProcessStartInfo("explorer.exe", " /select, " + raporPath));
            }
            Environment.Exit(0);
        }

        //Reads ekstre and formul file and creates rapor
        private static ExcelPackage CreateRapor(ExcelPackage ekstre, ExcelPackage formul)
        {
            Console.WriteLine("Rapor oluşturuluyor.");
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            ExcelWorksheet ekstreWorksheet = ekstre.Workbook.Worksheets[1];
            Console.WriteLine("Rapor oluşturuluyor..");
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            ExcelWorksheet formulWorksheet = formul.Workbook.Worksheets[1];
            Console.WriteLine("Rapor oluşturuluyor...");

            List<string> raporWorksheets = new List<string>
            {
                islemSonucuWorksheetName,
                pivotRaporWorksheetName
            };
            ExcelPackage rapor = excelHandler.CreateFile(raporPath, raporWorksheets);

            ExcelWorksheet raporWorksheet = rapor.Workbook.Worksheets[1];
            List<Report> reports = new List<Report>();

            if (ekstreWorksheet.Dimension == null)
            {
                Console.WriteLine("Worksheet boş.");
                Console.ReadKey();
                Environment.Exit(0);
            }
            for (int row = 1; row <= ekstreWorksheet.Dimension.End.Row; row++)
            {
                if (row == 1)
                {
                    SetRaporHeaders(raporWorksheet);
                }
                else
                {
                    Report report = GetRapor(excelHandler.GetRows(ekstreWorksheet, row), formulWorksheet, false);
                    SetRapor(raporWorksheet, report, row);
                    reports.Add(report);
                    Console.WriteLine("%" + ((float)row / ekstreWorksheet.Dimension.End.Row) * 100);
                    Console.SetCursorPosition(0, Console.CursorTop - 1);
                }
            }
            Console.WriteLine("Rapor oluşturuldu.");

            PivotTable(rapor, rapor.Workbook.Worksheets[1], rapor.Workbook.Worksheets[2]);
            try
            {
                rapor.Save();
                Console.WriteLine("Rapor kaydedildi.");
                return rapor;
            }
            catch (Exception e)
            {
                ExceptionMessage(e);
            }
            return null;
        }

        //Calculates fee using cell parameters and target formul sheet
        private static Report GetRapor(ExcelRange[] cells, ExcelWorksheet formulSheet, bool debug = false)
        {
            int siraNo = Convert.ToInt32(cells[0].Value);
            int adet = Convert.ToInt32(cells[1].Value);
            float kgDesi = Convert.ToSingle(cells[2].Value);
            string mesafe = cells[3].Value?.ToString();

            if (debug)
            {
                Console.WriteLine(siraNoHeader + " : " + siraNo);
            }
            float ucret = GetFee(formulSheet, adet, kgDesi, mesafe, debug);

            Report report = new Report(siraNo, adet, kgDesi, mesafe, ucret);
            return report;
        }

        //Set cells values of rapor
        private static void SetRapor(ExcelWorksheet raporWorksheet, Report rapor, int row)
        {
            raporWorksheet.Cells[row, 1].Value = rapor.siraNo;
            raporWorksheet.Cells[row, 2].Value = rapor.adet;
            raporWorksheet.Cells[row, 3].Value = rapor.kgDesi;
            raporWorksheet.Cells[row, 4].Value = rapor.mesafe;
            raporWorksheet.Cells[row, 5].Value = rapor.ucret;
        }

        //Set cells headers of rapor
        private static void SetRaporHeaders(ExcelWorksheet raporWorksheet)
        {
            raporWorksheet.Cells[1, 1].Value = siraNoHeader;
            raporWorksheet.Cells[1, 2].Value = adetHeader;
            raporWorksheet.Cells[1, 3].Value = kgDesiHeader;
            raporWorksheet.Cells[1, 4].Value = mesafeHeader;
            raporWorksheet.Cells[1, 5].Value = ucretHeader;
        }

        //Calculates fee using parameters and target formul sheet
        private static float GetFee(ExcelWorksheet formulSheet, int adet, float kgDesi, string mesafe, bool debug = false)
        {
            ExcelRange[] desiCells = excelHandler.GetColumns(formulSheet, 1);
            float desiMin = 0;
            float desiMax = 0;
            float ucret = 0;
            for (int i = 1; i < desiCells.Length; i++)
            {
                string desiRange = desiCells[i].Value?.ToString().Trim();

                if (desiRange == "Artan Her Desi İçin")
                {
                    ucret = Convert.ToSingle(formulSheet.Cells[i, formulSheet.Cells[1, 2].Value.ToString().Contains(mesafe) ? 2 : 3].Value) +
                        ((kgDesi - desiMax) * Convert.ToSingle(formulSheet.Cells[i + 1, formulSheet.Cells[1, 2].Value.ToString().Contains(mesafe) ? 2 : 3].Value));
                    ucret *= adet;
                    if (debug)
                    {
                        Console.WriteLine(adetHeader + " : " + adet);
                        Console.WriteLine(kgDesiHeader + " : " + kgDesi);
                        Console.WriteLine(mesafeHeader + " : " + mesafe);
                        Console.WriteLine("DESİ ÜCRET ARALIĞI : " + artanHeader);
                        Console.WriteLine(ucretHeader + " : " + ucret);
                    }
                    return ucret;
                }

                string[] desiSplited = desiRange.Split('-');
                desiMin = Convert.ToSingle(desiSplited[0]);
                desiMax = Convert.ToSingle(desiSplited[1]);

                if (kgDesi >= desiMin && kgDesi <= desiMax)
                {
                    ucret = Convert.ToSingle(formulSheet.Cells[i + 1, formulSheet.Cells[1, 2].Value.ToString().Contains(mesafe) ? 2 : 3].Value);
                    ucret *= adet;
                    if (debug)
                    {
                        Console.WriteLine(adetHeader + " : " + adet);
                        Console.WriteLine(kgDesiHeader + " : " + kgDesi);
                        Console.WriteLine(mesafeHeader + " : " + mesafe);
                        Console.WriteLine("DESİ ÜCRET ARALIĞI : " + desiMin + "-" + desiMax);
                        Console.WriteLine(ucretHeader + " : " + ucret);
                    }
                    return ucret;
                }
            }
            return 0;
        }

        //Created pivot table
        private static void PivotTable(ExcelPackage package, ExcelWorksheet worksheetData, ExcelWorksheet worksheetPivot)
        {
            Console.WriteLine("Pivot raporu oluşturuluyor...");
            var dataRange = worksheetData.Cells[worksheetData.Dimension.Address];

            //create the pivot table
            var pivotTable = worksheetPivot.PivotTables.Add(worksheetPivot.Cells["B2"], dataRange, "PivotRapor");

            //label field
            pivotTable.RowFields.Add(pivotTable.Fields["MESAFE"]);
            pivotTable.DataOnRows = false;

            //data fields
            var field = pivotTable.DataFields.Add(pivotTable.Fields["ADET"]);
            field.Name = "Kargo Adeti";
            field.Function = DataFieldFunctions.Count;

            Console.WriteLine("Pivot raporu oluşturuldu.");
        }

        private static void ExceptionMessage(Exception e)
        {
            Console.WriteLine(e);
            Console.ReadKey();
            Environment.Exit(0);
        }

        private static void FileCheck(ExcelPackage excelPackage, string path)
        {
            if (excelPackage == null)
            {
                Console.WriteLine("Dosya bulunamadı : " + path);
                Console.WriteLine("Lütfen dosyaları kontrol ediniz...");
                Console.ReadKey();
                Environment.Exit(0);
            }
        }
    }
}
