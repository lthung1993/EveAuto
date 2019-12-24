using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Automation
{
    public class AutomationBase
    {
        #region Khai báo thuộc tính, biến tĩnh, biến dùng chung
        protected static IWebDriver _webdriver;
        public IWebDriver driver { get; set; }
        public static DateTime testingDay { get; set; }

        public static string TestCaseID;
        public static string _log;

        public static string _reportFolder;
        public System.Diagnostics.Stopwatch TimeRun { get; set; }

        #endregion

        #region Build hàm SetUp - Teardown
        [SetUp]
        public void SetUp()
        {
            TimeRun = new System.Diagnostics.Stopwatch();
            TimeRun.Start();
            testingDay = DateTime.Now;


            if (_webdriver != null)
            {
                driver = _webdriver;
                return;
            }

            #region Config Webdriver
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            options.AddArguments("chrome.switches", "--disable-extensions --disable-extensions-file-access-check --disable-extensions-http-throttling");
            options.AddArgument("--dns-prefetch-disable");
            options.AddArgument("--disable-plugins");
            options.AddArgument("--disable-infobars");
            options.AddArgument("--disable-cache");
            options.AddArgument("--no-sandbox");

            foreach (var temp in Process.GetProcessesByName("chrome"))
                temp.Kill();
            this.driver = new ChromeDriver(options);
            #endregion
            if (_webdriver == null)
            {
                _webdriver = driver;
            }
        }

        [TearDown]
        public void TearDown()
        {
            TimeRun.Stop();
            //Gặp lỗi HTTP request hoặc Aw,snap thì xóa chrome chạy lại
            if (TestContext.CurrentContext.Result.Message != null)
            {
                if (TestContext.CurrentContext.Result.Message.Contains("The HTTP request to the remote WebDriver server for URL") ||
                    TestContext.CurrentContext.Result.Message.Contains("target window already closed"))
                {
                    QuitDriver();
                }

            }
            KillSpecificExcelFileProcess();
            ReportExcel();
            Console.WriteLine("Total time run :   {0}", TimeRun.Elapsed.ToString(@"hh\:mm\:ss"));
        }
        #endregion

        #region Viết lại một số hàm cơ bản
        public IWebElement FindWebElement(By by, IWebDriver webDriver = null)
        {
            if (driver == null)
            {
                IWebElement element = AutomationHelper.WaitElement(driver, by);
                if (element == null)
                    return null;
                else
                    return element;
            }
            else
            {
                IWebElement element = AutomationHelper.WaitElement(webDriver, by);
                if (element == null)
                    return null;
                else
                    return element;
            }
        }
        protected List<IWebElement> FindWebElements(By by)
        {
            List<IWebElement> listResult = new List<IWebElement>();
            ReadOnlyCollection<IWebElement> elements = AutomationHelper.WaitElements(driver, by);
            if (elements == null) return null;
            foreach (IWebElement element in elements)
            {
                listResult.Add(element);
            }
            return listResult;
        }

        #endregion

        #region Đọc file excel
        public List<object> ReadDataSourceXlsx(string fileName, int sheetIndex, int rowStart, int columnStart, int rowIndex = 0)
        {
            string filePath = "";
            filePath = AppDomain.CurrentDomain.BaseDirectory + "DataSource\\" + fileName;
            List<object> lisRowData = new List<object>();

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(filePath);
                Worksheet ws = wb.Worksheets[sheetIndex + 1];
                Microsoft.Office.Interop.Excel.Range range = ws.UsedRange;
                Microsoft.Office.Interop.Excel.Range _range;

                int lastUsedRow = 0;
                int lastUsedColumn = 0;

                lastUsedRow = ws.Cells.Find("*", System.Reflection.Missing.Value,
                                 System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                 Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                 false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                lastUsedColumn = ws.Cells.Find("*", System.Reflection.Missing.Value,
                                 System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                 Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                 false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                if (rowIndex != 0)
                {
                    rowStart = rowIndex;
                    lastUsedRow = rowIndex;
                    lastUsedColumn = ws.Rows[rowIndex].Cells.Length();
                }

                for (int row = rowStart; row <= lastUsedRow; row++)
                {
                    List<string> listValuesCell = new List<string>();
                    for (int col = columnStart; col <= lastUsedColumn; col++)
                    {

                        _range = (Microsoft.Office.Interop.Excel.Range)ws.Cells[row, col];
                        string cellValue = _range.Text.Trim().ToString() ?? "";
                        listValuesCell.Add(cellValue);

                    }
                    lisRowData.Add(listValuesCell.ToArray());
                }
                excel.Quit();
            }
            catch (Exception e)
            {
                excel.Quit();
            }
            return lisRowData;
        }
        #endregion

        #region Tạo báo cáo
        private string GetReportFolder()
        {
            //Mỗi tuần chạy => sinh ra một thư mục chung tuần hiện tại
            //Mỗi lần chạy => Tăng dần đều (lên 1 đơn vị) cho cái tên thư mục

            //VD: Tuần 50 - Chạy lần đầu tiên => Debug/Report/Week50
            //VD: Tuần 50 - Chạy lần đầu 2 => ../Debug/Report/Week50_1
            //VD: Tuần 50 - Chạy lần đầu 3 =>../Debug/Report/Week50_2
            if (!string.IsNullOrEmpty(_reportFolder))
            {
                return _reportFolder;
            }


            string rootPath = AppDomain.CurrentDomain.BaseDirectory + "Report";

            string currentFolderReport = rootPath + "\\Week" + CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(DateTime.Today, CultureInfo.CurrentUICulture.DateTimeFormat.CalendarWeekRule, CultureInfo.CurrentUICulture.DateTimeFormat.FirstDayOfWeek).ToString();

            int i = 0;
            while (true)
            {
                if (!Directory.Exists(currentFolderReport))
                {
                    Directory.CreateDirectory(currentFolderReport);
                    _reportFolder = currentFolderReport;
                    return currentFolderReport;
                }
                else
                {
                    i++;
                    currentFolderReport = Regex.Split(currentFolderReport, "_").ToList()[0];
                    currentFolderReport = currentFolderReport + "_" + i;
                }
            }
        }
        private void AllBorders(Microsoft.Office.Interop.Excel.Borders _boder)
        {
            _boder[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            _boder[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            _boder[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            _boder[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            _boder.Color = System.Drawing.Color.Black;
        }
        private void ReportExcel()
        {
            var excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;

            string filePath = GetReportFolder() + "\\Report.xlsx";
            var Page = Regex.Split(this.ToString(), @"\.").ToList();
            if (!File.Exists(filePath))
            {
                excel.Application.Workbooks.Add(Type.Missing);

                #region Summary Report
                excel.Range["B3:D5"].Merge(Type.Missing);
                excel.Range["B3:D5"].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#54FF9F");
                excel.Range["B3:D5"].Font.Bold = true;
                excel.Range["B3"].Value = "Summary";
                excel.Range["B3"].Font.Size = 20;
                excel.Range["B3"].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Range["B3"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Range["B6"].Value = "Passed";
                excel.Range["C6"].Value = "Failed";
                excel.Range["D6"].Value = "Error";
                excel.Range["B7"].Formula = "=COUNTIF(F:F,\"Passed\")";
                excel.Range["C7"].Value = "=COUNTIF(F:F,\"Failed\")";
                excel.Range["D7"].Value = "=COUNTIF(F:F,\"Error\")";
                AllBorders(excel.Range["B3:D7"].Borders);
                #endregion

                #region Testting Information
                excel.Range["B11:C13"].Merge();
                excel.Range["B11:C13"].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#54FF9F");
                excel.Range["B11:C13"].Font.Bold = true;
                excel.Range["B11"].Value = "Testting Information";
                excel.Range["B11"].Font.Size = 20;
                excel.Range["B11"].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Range["B11"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Range["B14"].Value = "Total Time Run";
                excel.Range["B15"].Value = "Start time";
                excel.Range["B16"].Value = "Stop time";
                //Data
                excel.Range["C14"].Formula = "=SUM(H:H)";
                excel.Range["C14"].NumberFormat = "hh:mm:ss";
                excel.Range["C15"].Value = testingDay.ToString("dd/MM/yyyy HH:mm:ss");
                AllBorders(excel.Range["B11:C16"].Borders);


                #endregion

                #region Chart

                var ws = (Microsoft.Office.Interop.Excel.Worksheet)excel.Worksheets.get_Item(1);
                var chart = (Microsoft.Office.Interop.Excel.ChartObjects)ws.ChartObjects(Type.Missing);
                var mychart = (Microsoft.Office.Interop.Excel.ChartObject)chart.Add(630, 10, 250, 250);
                var chartPage = (Microsoft.Office.Interop.Excel.Chart)mychart.Chart;


                var seriesCollection = chartPage.SeriesCollection();
                var series = seriesCollection.NewSeries();
                series.XValues = ws.Range["B6", "D6"];
                series.Values = ws.Range["B7", "D7"];

                chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie;
                chartPage.ApplyLayout(6);
                chartPage.ChartTitle.Text = "Summary";
                chartPage.ChartTitle.Font.Size = 10;
                ((Microsoft.Office.Interop.Excel.LegendEntry)chartPage.Legend.LegendEntries(1)).LegendKey.Interior.Color = (int)Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlue;
                ((Microsoft.Office.Interop.Excel.LegendEntry)chartPage.Legend.LegendEntries(2)).LegendKey.Interior.Color = (int)Microsoft.Office.Interop.Excel.XlRgbColor.rgbRed;
                ((Microsoft.Office.Interop.Excel.LegendEntry)chartPage.Legend.LegendEntries(3)).LegendKey.Interior.Color = (int)Microsoft.Office.Interop.Excel.XlRgbColor.rgbLightSalmon;


                #endregion

                #region Header
                excel.Range["B21"].Value = "TestCase ID";
                excel.Range["C21"].Value = "Page";
                excel.Range["D21"].Value = "TestCase Name";
                excel.Range["E21"].Value = "URL";
                excel.Range["F21"].Value = "KetQua";
                excel.Range["G21"].Value = "Log";
                excel.Range["H21"].Value = "Thời gian test";
                excel.Range["B21:H21"].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#54FF9F");
                excel.Range["B21:H21"].Font.Size = 14;
                AllBorders(excel.Range["B21:H21"].Borders);

                #endregion

                #region Data
                //Chạy lần đầu tiên => KQ sẽ thêm dòng 22
                //Lưu file lại này => Tồn tại file
                //Lấy được dòng cuối cùng được sử dụng
                //Thêm dữ liệu vào dòng cuối cùng + 1
                //Lưu
                excel.Range["B22"].Value = TestCaseID;
                //Tên màn hình:
                //Truyền vào hoặc lấy dữ liệu từ file data test
                //Sử dụng namespace
                excel.Range["C22"].Value = Page[1];
                excel.Range["D22"].Value = TestContext.CurrentContext.Test.MethodName;
                excel.Range["E22"].Value = driver.Url.ToString();
                excel.Range["F22"].Value = TestContext.CurrentContext.Result.Outcome.Status.ToString();
                //Log
                //Framework tự sinh log => Override một số hàm, khi mà chạy những hàm này => sinh ra 1 file log
                //Chúng ta tự thêm log
                excel.Range["G22"].Value = _log.ToString();
                excel.Range["H22"].NumberFormat = "hh:mm:ss";
                excel.Range["H22"].Value = TimeRun.Elapsed.ToString(@"hh\:mm\:ss");
                AllBorders(excel.Range["B22:H22"].Borders);
                #endregion

                excel.Range["C16"].Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                excel.ActiveWorkbook.SaveCopyAs(filePath);
            }
            else
            {
                #region Data
                //Lấy dòng được sử dụng cuối cùng: n
                //Thêm dòng n+1
                // Lấy dòng được sử dụng cuối cùng như thế nào???
                Microsoft.Office.Interop.Excel.Workbook wbv = excel.Workbooks.Open(filePath);
                Microsoft.Office.Interop.Excel.Worksheet ws = wbv.Worksheets.get_Item(1);
                int lastRow = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;

                int presentCell = lastRow + 1;
                excel.Range["B" + presentCell].Value = TestCaseID;
                //Tên màn hình:
                //Truyền vào hoặc lấy dữ liệu từ file data test
                //Sử dụng namespace
                excel.Range["C" + presentCell].Value = Page[1];
                excel.Range["D" + presentCell].Value = TestContext.CurrentContext.Test.MethodName;
                excel.Range["E" + presentCell].Value = driver.Url.ToString();
                excel.Range["F" + presentCell].Value = TestContext.CurrentContext.Result.Outcome.Status.ToString();
                //Log
                //Framework tự sinh log => Override một số hàm, khi mà chạy những hàm này => sinh ra 1 file log
                //Chúng ta tự thêm log
                excel.Range["G" + presentCell].Value = _log.ToString();
                excel.Range["H" + presentCell].NumberFormat = "hh:mm:ss";
                excel.Range["H" + presentCell].Value = TimeRun.Elapsed.ToString(@"hh\:mm\:ss");
                AllBorders(excel.Range["B" + presentCell + ":H" + presentCell].Borders);
                #endregion


                excel.Range["C16"].Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                excel.ActiveWorkbook.Save();
            }
            excel.ActiveWorkbook.Saved = true;
            excel.ActiveWorkbook.Close();
            excel.Quit();
        }
        #endregion

        #region Hàm hỗ trợ
        private void KillSpecificExcelFileProcess()
        {
            foreach (Process clsProcess in Process.GetProcesses())
                if (clsProcess.ProcessName.Equals("EXCEL"))  //Process Excel?
                    clsProcess.Kill();
        }
        protected void QuitDriver()
        {
            try
            {
                _webdriver = null;
                driver.Quit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
    }
}
