using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eplus_example
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage package = new ExcelPackage(new FileInfo("Academy Awards.xlsx"));
            ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault();
            IEnumerable<AcademyAward> awards = PopulateAwards(workSheet, true);
            Console.WriteLine("Awards count: " + awards.Count());
            Console.ReadLine();
        }

        /// <summary>
        /// Populate award objects from spreadsheet
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="firstRowHeader"></param>
        /// <returns></returns>
        static IEnumerable<AcademyAward> PopulateAwards(ExcelWorksheet workSheet, bool firstRowHeader)
        {
            IList<AcademyAward> awards = new List<AcademyAward>();

            if (workSheet != null)
            {
                Dictionary<string, int> header = new Dictionary<string,int>();

                for (int rowIndex = workSheet.Dimension.Start.Row; rowIndex <= workSheet.Dimension.End.Row; rowIndex++)
                {
                    //Assume the first row is the header.  Then use the column match ups by name to determine the index.
                    //This will allow you to have the order of the columns change without any affect.
                    
                    if (rowIndex == 1 && firstRowHeader)
                    {
                        header = ExcelHelper.GetExcelHeader(workSheet, rowIndex);
                    }
                    else
                    {
                        awards.Add(new AcademyAward{
                            Year = ExcelHelper.ParseWorksheetValue(workSheet, header, rowIndex, "Year"),
                            Category = ExcelHelper.ParseWorksheetValue(workSheet, header, rowIndex, "Category"),
                            Nominee = ExcelHelper.ParseWorksheetValue(workSheet, header, rowIndex, "Nominee"),
                            AdditionalInfo = ExcelHelper.ParseWorksheetValue(workSheet, header, rowIndex, "AdditionalInfo"),
                            Won = ExcelHelper.ParseWorksheetValue(workSheet, header, rowIndex, "Won?")
                        });

                    }
                }
            }

            return awards;
        }
    }
}
