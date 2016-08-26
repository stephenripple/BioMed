using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;
using System.IO;

namespace BioMed.Parsers
{
    public class ExcelParser
    {
        private FileInfo ExcelFileInfo;

        public ExcelParser(string FilePath)
        {
            //Add validation for file
            ExcelFileInfo = new FileInfo(FilePath);
        }

        public ExcelParser(FileInfo File)
        {
            ExcelFileInfo = File;
        }
        public string toXML()
        {
            using (var excelPackage = new ExcelPackage(ExcelFileInfo))
            {
                foreach (var worksheet in excelPackage.Workbook.Worksheets)
                {
                    var xml = worksheet.Cells.Text;
                    return xml.ToString();
                }

            }
            return string.Empty;
        }

        public List<string> TabToList(string tab)
        {
            using (var excelPackage = new ExcelPackage(ExcelFileInfo))
            {
                var worksheet = excelPackage.Workbook.Worksheets[tab];
                var columncount = worksheet.Dimension.Columns;
                var rowcount = worksheet.Dimension.Rows;
            }
            return null;
        }
    }    
}
