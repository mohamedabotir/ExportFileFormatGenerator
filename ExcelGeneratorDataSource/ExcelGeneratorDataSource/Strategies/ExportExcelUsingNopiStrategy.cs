using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
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
    public class ExportExcelUsingNopiStrategy : ExportStrategy
    {
        public string Export(object datasource)
        {
        DataTable data = (DataTable)datasource;

            // Create the Excel file
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            // Create header row
            IRow headerRow = sheet.CreateRow(0);
            for (int i = 0; i < data.Columns.Count; i++)
            {
                headerRow.CreateCell(i).SetCellValue(data.Columns[i].ColumnName);
            }

            // Create data rows
            for (int i = 0; i < data.Rows.Count; i++)
            {
                IRow dataRow = sheet.CreateRow(i + 1);
                for (int j = 0; j < data.Columns.Count; j++)
                {
                    dataRow.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                }
            }

            // Save the Excel file
            string serverPath = @"E:\Path\To\Save\Excel\File.xls";
            createPathIfNotExists(serverPath);
            using (FileStream stream = new FileStream(serverPath, FileMode.Create,FileAccess.ReadWrite))
            {
                workbook.Write(stream);
            }

          
            return serverPath;
        }

        private  void createPathIfNotExists(string path)
        {

            if (!File.Exists(path)) 
                Directory.CreateDirectory(path);
        }

    }
}
