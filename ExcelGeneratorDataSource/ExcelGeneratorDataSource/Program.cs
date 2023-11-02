using ExcelGeneratorDataSource.Strategies;
using System.Data;

namespace ExcelGeneratorDataSource
{
    internal class Program
    {
        static void Main(string[] args)
        {


            //string exportDataSourceUsingEPPlus = new ExportExcelUsingEPPlusStrategy().Export(GetDataFromDatabase());
            string exportDataSourceUsingNopi = new ExportExcelUsingNopiStrategy().Export(GetDataFromDatabase());


        }

        static DataTable GetDataFromDatabase()
        {
            // Replace this with your actual data retrieval logic
            DataTable data = new DataTable();
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Age", typeof(string));

            data.Rows.Add("Abdullah maher", 26);
            data.Rows.Add("Abotir", 23);

            return data;
        }
    }
}