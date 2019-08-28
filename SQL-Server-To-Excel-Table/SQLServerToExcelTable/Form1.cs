using Syncfusion.XlsIO;
using System;
using System.Windows.Forms;

namespace SQLServerToExcelTable
{
    public partial class Form1 : Form
    {
        public static string DataPathBase = @"..\..\Data\";
        public static string DataPathOutput = @"..\..\Output\";
        public Form1()
        {
            InitializeComponent();
        }

        private void refreshExcelTable_Click(object sender, EventArgs e)
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Excel2016;

            //Create a new workbook 
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet sheet = workbook.Worksheets[0];

            if (sheet.ListObjects.Count == 0)
            {
                //Establishing the connection in the worksheet 
                string connectionString = "Server=myServerAddress;Database=myDataBase;User Id=myUsername;Password = myPassword";

                string query = "SELECT * FROM Employee_Details";

                IConnection connection = workbook.Connections.Add("SQLConnection", "Sample connection with SQL Server", connectionString, query, ExcelCommandType.Sql);

                //Create Excel table from external connection. 
                sheet.ListObjects.AddEx(ExcelListObjectSourceType.SrcQuery, connection, sheet.Range["A1"]);
            }

            //Refresh Excel table to get updated values from database 
            sheet.ListObjects[0].Refresh();

            sheet.UsedRange.AutofitColumns();

            //Save the file in the given path
            string outputPath = DataPathOutput + "CreateExcelFile.xlsx";
            workbook.SaveAs(outputPath);

            workbook.Close();
            excelEngine.Dispose();
        }
    }
}
