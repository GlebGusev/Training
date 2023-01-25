using Training.Helpers;

namespace Training
{
    public class PlanTests
    {
        private string xlsPath = @"C:\\Users\\AndriusBogda\\Downloads\\ImportFile.xls";
        private string csvPath = @"C:\\Users\\AndriusBogda\\Downloads\\ExportFile.csv";

        [Test]
        public void CreateCSVFromXlsTest()
        {
            CsvHelper.WriteTo(
                csvPath,
                ExcelHelper.Deserialize(xlsPath, true));
            
            //what you code is doing is simulating simple SaveAs. There is basic code that allows you to do so
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(@"C:\\Users\\AndriusBogda\\Downloads\\ImportFile.xls");
            wb.SaveAs(@"C:\\Users\\AndriusBogda\\Downloads\\ExportFile.csv", Excel.XlFileFormat.xlCSVWindows);
            wb.Close(false);
            excel.Quit();

            Assert.IsTrue(File.Exists(csvPath));
        }

        [Test]
        public void ValidateCSV()
        {
            var csv = ExcelHelper.Deserialize(csvPath);
            var xls = ExcelHelper.Deserialize(xlsPath, true);

            for (int i = 0; i < xls.Length; i++)
            {
                for (int j = 0; j < xls[i].Length; j++)
                {
                    Assert.AreEqual(xls[i][j], csv[j][i].Replace(",", ""));
                }
            }
        }
    }
}
