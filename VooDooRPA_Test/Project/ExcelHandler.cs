using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VooDooRPA_Project
{
    internal class ExcelHandler
    {
        //Creates excel file defined path
        public ExcelPackage CreateFile(string path, List<string> worksheets)
        {
            try
            {
                FileInfo newFile = new FileInfo(path);
                if (newFile.Exists)
                {
                    newFile.Delete();
                    newFile = new FileInfo(path);
                }
                ExcelPackage package = new ExcelPackage(newFile);

                for (int i = 0; i < worksheets.Count; i++)
                {
                    package.Workbook.Worksheets.Add(worksheets[i]);
                }
                package.Save();
                Console.WriteLine("Dosya oluşturuldu : " + path);
                return package;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;

            }
        }

        //Returns worksheets of excel file from path
        public ExcelPackage GetExcelPackage(string FilePath)
        {
            FileInfo existingFile = new FileInfo(FilePath);

            if (existingFile.Exists)
            {
                Console.WriteLine("Dosya bulundu : " + existingFile.Name);

                ExcelPackage package = new ExcelPackage(existingFile);
                return package;
            }
            else
            {
                return null;
            }
        }

        //Returns cell array in row
        public ExcelRange[] GetRows(ExcelWorksheet worksheet, int row)
        {
            ExcelRange[] cells = new ExcelRange[worksheet.Dimension.End.Column];
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                cells[(col - 1)] = worksheet.Cells[row, col];
            }
            return cells;
        }

        //Returns cell array in column
        public ExcelRange[] GetColumns(ExcelWorksheet worksheet, int column)
        {
            ExcelRange[] cells = new ExcelRange[worksheet.Dimension.End.Row];
            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                cells[(row - 1)] = worksheet.Cells[row, column];
            }
            return cells;
        }
    }
}
