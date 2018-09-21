using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Windows.Media;

namespace OpenSoftTest
{
    class ExcelScanner: IDisposable
    {
        Application excelApplication;

        public ExcelScanner()
        {
            excelApplication = new Application();
        }

        public void paintWords(string filePath, string word, XlRgbColor color)
        {
            Workbook workBook = excelApplication.Workbooks.Open(filePath);
            int sheets = workBook.Sheets.Count;

            try
            {
                for (int i = 1; i < sheets + 1; i++)
                {
                    Worksheet sheet = (Worksheet)excelApplication.Sheets[i];
                    if (sheet.UsedRange.Count > 1)
                    {
                        object[,] valueArray = (object[,])sheet.UsedRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
                        for (int h = 1; h < valueArray.GetLength(0) + 1; h++)
                            for (int m = 1; m < valueArray.GetLength(1) + 1; m++)
                                if (valueArray[h, m] != null && valueArray[h, m].ToString() == word) sheet.Cells[h, m].Interior.Color = color;

                    }
                    else
                    {
                        string value = sheet.UsedRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
                        if (value == word) sheet.Cells[sheet.UsedRange.Row, sheet.UsedRange.Column].Interior.Color = XlRgbColor.rgbRed;
                    }
                }
                workBook.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            }
            finally
            {
                workBook.Close(false, filePath, null);
                Marshal.ReleaseComObject(workBook);
            }
        }

        public void Dispose()
        {
            Marshal.ReleaseComObject(excelApplication);
        }


    }
}
