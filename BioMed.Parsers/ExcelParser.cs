using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;
using System.IO;
using BioMed.Data.Models;

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

        public List<BioMedProduct> TabToList(string tab)
        {
            var ret = new List<BioMedProduct>();
            using (var excelPackage = new ExcelPackage(ExcelFileInfo))
            {
                var worksheet = excelPackage.Workbook.Worksheets[tab];
                var columncount = worksheet.Dimension.Columns;
                var rowcount = worksheet.Dimension.Rows;
                for  (int x = 2; x <= rowcount; x++)
                {
                    var product = new BioMedProduct();
                    product.Make = worksheet.Cells[x, 1].Value.ToString();
                    product.Model = worksheet.Cells[x, 2].Value.ToString();
                    product.Specs = new List<BioMedProductSpec>();
                    for (int y = 3; y <= columncount; y++)
                    {
                        var spec = new BioMedProductSpec();
                        spec.SpecTitle = worksheet.Cells[1, y].Value.ToString();
                        spec.Spec = worksheet.Cells[x, y].Value.ToString();
                        product.Specs.Add(spec);
                    }
                    ret.Add(product);
                }
            }
            return ret;
        }

        public List<BioMedProduct> TabsToList()
        {
            var ret = new List<BioMedProduct>();
            using (var excelPackage = new ExcelPackage(ExcelFileInfo))
            {
                foreach (var worksheet in excelPackage.Workbook.Worksheets)
                {
                    if (worksheet.Name != "Patient Monitoring")
                        ret.Concat(TabToList(worksheet.Name));
                }
            }
            return ret;
        }
    }    
}
