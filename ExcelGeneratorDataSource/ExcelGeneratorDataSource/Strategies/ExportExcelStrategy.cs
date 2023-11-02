using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace ExcelGeneratorDataSource.Strategies
{
    // Install required  Package  Install-Package EPPlus required license in commercial situation or NPOI free to use appache license  Install-Package NPOI -Version 2.6.2
    public class ExportExcelUsingEPPlusStrategy : ExportStrategy
    {
        public string Export(object datasource)
        {
            DataTable data = (DataTable)datasource;

            // Create the Excel file
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells.LoadFromDataTable(data, true);

                // Save the Excel file
                string serverPath = @"C:\Path\To\Save\Excel\File.xls";
                FileInfo file = new FileInfo(serverPath);
                package.SaveAs(file);

                // Provide download link to the client
                string downloadLink = serverPath;

                // Return the download link to the client
                return downloadLink;
            }


        }
    }
}
