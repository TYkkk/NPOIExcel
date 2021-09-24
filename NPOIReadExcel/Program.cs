using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOIReadExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    //ReadExcelData(args[i]);
                    WriteTxtToExcel(args[i]);
                }
            }

            Console.WriteLine("End");

            Console.ReadKey();
        }

        private static void ReadExcelData(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return;
            }

            IWorkbook wk = null;

            string extension = Path.GetExtension(filePath);
            string savePath = Path.GetDirectoryName(filePath);
            try
            {
                FileStream fs = File.OpenRead(filePath);

                int sheetCount = 0;

                if (extension.Equals(".xls"))
                {
                    wk = new HSSFWorkbook(fs);
                    sheetCount = (wk as HSSFWorkbook).NumberOfSheets;
                }
                else
                {
                    wk = new XSSFWorkbook(fs);
                    sheetCount = (wk as XSSFWorkbook).NumberOfSheets;
                }

                fs.Close();
                fs.Dispose();

                for (int n = 0; n < wk.NumberOfSheets; n++)
                {
                    ISheet sheet = wk.GetSheetAt(n);

                    if (sheet.SheetName[0] != '#')
                    {
                        continue;
                    }

                    StringBuilder result = new StringBuilder();

                    for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            string content = row.Cells[j].ToString();

                            if (row.Cells[j].CellType == CellType.Formula)
                            {
                                row.Cells[j].SetCellType(CellType.String);
                                content = row.Cells[j].StringCellValue;
                            }

                            content = content.Replace("\"", "\"\"");

                            if (content.Contains(',') || content.Contains('"') || content.Contains('\r') || content.Contains('\n'))
                            {
                                content = string.Format("\"{0}\"", content);
                            }

                            result.Append(content);

                            if (j != row.Cells.Count - 1)
                            {
                                result.Append(",");
                            }
                        }

                        if (i != sheet.LastRowNum)
                        {
                            result.Append("\n");
                        }
                    }

                    if (!Directory.Exists("./Result"))
                    {
                        Directory.CreateDirectory("./Result");
                    }

                    File.WriteAllText($"./Result/{sheet.SheetName.Remove(0, 1)}.txt", result.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static void ReadCSVData(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return;
            }

            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            StreamReader sr = new StreamReader(fs);

            var csv = new CsvHelper.CsvReader(sr, CultureInfo.InvariantCulture);

            var result = csv.GetRecords<Student>();

            foreach (var child in result)
            {
                Console.WriteLine(child.Name + "&" + child.Sex + "&" + child.Age + "&" + child.Desc);
            }

            csv.Dispose();
            sr.Dispose();
            fs.Dispose();
        }

        private static void WriteTxtToExcel(string filePath)
        {
            var dataStr = File.ReadAllText(filePath);

            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("sheet0");

            StringReader sr = new StringReader(dataStr);
            string temp = "";

            int rowIndex = 0;

            while ((temp = sr.ReadLine()) != null)
            {
                string[] data = temp.Split(',');

                IRow row = sheet.CreateRow(rowIndex);

                for (int i = 0; i < 4; i++)
                {
                    if (!string.IsNullOrEmpty(data[i]))
                    {
                        ICell cell = row.CreateCell(i);
                        cell.SetCellValue(data[i]);
                    }
                }

                rowIndex++;
            }


            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();

            using (FileStream fs = new FileStream(filePath.Replace(".txt", ".xlsx"), FileMode.OpenOrCreate, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }
    }
}

public class Student
{
    public string Name;
    public string Sex;
    public int Age;
    public string Desc;
}