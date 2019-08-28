using System;
using Syncfusion.XlsIO;
using System.Windows.Forms;
using Syncfusion.XlsIO.Implementation;

namespace QueryTableParameters
{
    public partial class Form1 : Form
    {
        public static string DataPathBase = @"..\..\Data\";
        public static string DataPathOutput = @"..\..\Output\";
        public Form1()
        {
            InitializeComponent();
        }

        private void constantParam_Click(object sender, EventArgs e)
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Open(DataPathBase + "Template.xlsx");

            IWorksheet worksheet = workbook.Worksheets[0];
            QueryTableImpl queryTable = worksheet.ListObjects[0].QueryTable;

            string query = "select * from Employee_Details WHERE Emp_Age > ? And Country = ?;";
            queryTable.CommandText = query;

            IParameter constParam1 = queryTable.Parameters.Add("parameter1", ExcelParameterDataType.ParamTypeInteger);
            constParam1.SetParam(ExcelParameterType.Constant, 26);

            IParameter constParam2 = queryTable.Parameters.Add("parameter2", ExcelParameterDataType.ParamTypeInteger);
            constParam2.SetParam(ExcelParameterType.Constant, "Argentina");

            worksheet.ListObjects[0].Refresh();

            string outputPath = DataPathOutput + "ConstantParameter.xlsx";
            workbook.SaveAs(outputPath);
            workbook.Close();
            excelEngine.Dispose();
            System.Diagnostics.Process.Start(outputPath);
        }

        private void rangeParam_Click(object sender, EventArgs e)
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Open(DataPathBase + "Template.xlsx");

            IWorksheet worksheet = workbook.Worksheets[0];
            QueryTableImpl queryTable = worksheet.ListObjects[0].QueryTable;

            string query = "select * from Employee_Details WHERE Emp_Age > ?;";
            queryTable.CommandText = query;

            IParameter rangeParam = queryTable.Parameters.Add("RangeParameter", ExcelParameterDataType.ParamTypeInteger);
            rangeParam.SetParam(ExcelParameterType.Range, worksheet.Range["H1"]);
           
            worksheet.ListObjects[0].Refresh();

            string outputPath = DataPathOutput + "RangeParameter.xlsx";
            workbook.SaveAs(outputPath);
            workbook.Close();
            excelEngine.Dispose();
            System.Diagnostics.Process.Start(outputPath);
        }

        private void promptParam_Click(object sender, EventArgs e)
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Open(DataPathBase + "Template.xlsx");

            IWorksheet worksheet = workbook.Worksheets[0];
            QueryTableImpl queryTable = worksheet.ListObjects[0].QueryTable;

            string query = "select * from Employee_Details WHERE Emp_Age < ? AND Country = ?;";
            queryTable.CommandText = query;

            IParameter promptParam1 = queryTable.Parameters.Add("PromptParameter1", ExcelParameterDataType.ParamTypeInteger);
            promptParam1.SetParam(ExcelParameterType.Prompt, "PromptParameter1");
            promptParam1.Prompt += new PromptEventHandler(SetPromptParameter1);

            IParameter promptParam2 = queryTable.Parameters.Add("PrromptParameter2", ExcelParameterDataType.ParamTypeInteger);
            promptParam2.SetParam(ExcelParameterType.Prompt, "PromptParameter2");
            promptParam2.Prompt += new PromptEventHandler(SetPromptParameter2);

            worksheet.ListObjects[0].Refresh();

            string outputPath = DataPathOutput + "PromptParameter.xlsx";
            workbook.SaveAs(outputPath);
            workbook.Close();
            excelEngine.Dispose();
            System.Diagnostics.Process.Start(outputPath);
        }

        private void SetPromptParameter1(object sender, PromptEventArgs args)
        {
            args.Value = 28;
        }

        private void SetPromptParameter2(object sender, PromptEventArgs args)
        {
            args.Value = "Argentina";
        }
    }
}
