using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataFromExcelFile
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> excelData = new List<string>();

            byte[] bin = File.ReadAllBytes(@"C:\Users\mohamed.elkabir\Source\Repos\ImportDataFromExcelFile\ImportDataFromExcelFile\CSV_Signalis.xlsx");
            string stringContent = string.Empty;
            using (MemoryStream stream = new MemoryStream(bin))
           using(ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                var workSheet = excelPackage.Workbook.Worksheets.SingleOrDefault(x => x.Name == "CSV");
                // Parcours des lignes
                for (int i= workSheet.Dimension.Start.Row; i <= workSheet.Dimension.End.Row; i++)
                {
                    // Parcours des colonnes
                    for(int j = workSheet.Dimension.Start.Column; j <= workSheet.Dimension.End.Column; j++)
                    {
                        if (workSheet.Cells[i, j].Value != null)
                        {
                            excelData.Add(workSheet.Cells[i, j].Value.ToString());

                        }
                        else
                        {
                            excelData.Add(String.Empty);
                        }
                        
                        

                    }

                    stringContent += string.Join(";", excelData.ToArray()) + Environment.NewLine;
                    excelData.Clear();
                }
            }

            if(!String.IsNullOrEmpty(stringContent))
            {
                using(StreamWriter sw = new StreamWriter("Monficher.csv"))
                {
                    sw.Write(stringContent);
                }
            }

            Console.WriteLine("{0}", stringContent);
            Console.ReadLine();
        }
    }
}
