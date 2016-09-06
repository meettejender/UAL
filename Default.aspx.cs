using System;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using WebChart;
using System.Drawing;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Specialized;
using System.Reflection;
using System.Diagnostics;

public partial class _Default : System.Web.UI.Page
{
    private string path1 = @"D:\OTA DashBoard_Pricing_07252016.xlsx";
    //private string path2 = @"D:\TestData.xlsx";
    //private string path3 = @"D:\Book1.xlsx";
    private NameValueCollection nameValue = new NameValueCollection();


    protected void Page_Load(object sender, EventArgs e)
    {
        //loadNameValueCollection();
        //PlotChart();
        LoadExcel2();
        //example();
        //DataTable dt = ConvertToDataTable(path1);
    }

    private void loadNameValueCollection()
    {
        string[] values = null;
        nameValue.Add("Very High", "80");
        nameValue.Add("High", "60");
        nameValue.Add("medium", "50");
        nameValue.Add("Pass", "40");
        foreach (string key in nameValue.Keys)
        {
            values = nameValue.GetValues(key);
            foreach (string value in values)
            {
                //MessageBox.Show(key + " - " + value);
            }
        } 
    }

    private void example()
    {

        DataTable dt = new DataTable("Chart");
        dt.Columns.Add("IntValue", typeof(int));
        dt.Columns.Add("StringValue", typeof(string));
        dt.Rows.Add(1, "ONE");
        dt.Rows.Add(1, "ONE");
        dt.Rows.Add(1, "ONE");
        dt.Rows.Add(2, "TWO");
        dt.Rows.Add(2, "TWO");
   //     DataView view = new DataView(dt);
   //DataTable distinctValues = view.to .ToTable(true, "StringValue");
        //DataView view = new DataView(dt);
        //DataTable distinctValues = new DataTable();
        //distinctValues = view.ToTable(true, "StringValue");

        //var x = (from r in dt.AsEnumerable()
                 //select r["IntValue"]).Distinct().ToList();
        //var distinctWebsites = dt.AsEnumerable()
        //            .Select(s => new
        //            {
        //                id = s.Field<string>("StringValue"),
        //            })
        //            .Distinct().ToList();
        //var distinctValidations = dt.AsEnumerable()
        //    .Select(s => new
        //    {
        //        id = s.Field<string>("StringValue"),
        //    })
        //    .Distinct().Count();
        //var myResult = dt.AsEnumerable().Select(c => (DataRow)c["StringValue"]).Distinct().ToList();
        var result = dt.AsEnumerable()
               .GroupBy(r => r.Field<string>("StringValue"))
               .Select(r => new
               {
                   Str = r.Key,
                   Count = r.Count()
               });

       foreach (var item in result)
{
    string str = item.Str;
    int i = item.Count;
    //Console.WriteLine($"{item.Str} : {item.Count}");
}
    }

    private void PlotChart()
    {
        // Preparing Data Source For Chart Control
        DataTable dt = new DataTable("Chart");
        DataColumn dc = new DataColumn("Website", typeof(string));
        DataColumn dc1 = new DataColumn("Validation", typeof(int));
        dt.Columns.Add(dc);
        dt.Columns.Add(dc1);
        DataRow dr = dt.NewRow();
        dr[0] = "www.united.com";
        dr[1] = 184;
        dt.Rows.Add(dr);
        DataRow dr1 = dt.NewRow();
        dr1[0] = "www.orbits.com";
        dr1[1] = 9;
        dt.Rows.Add(dr1);
        //Chart type you can change chart type here ex. pie chart,circle chart

        LineChart chart = new LineChart();//Class instance for LineChart
        chart.Fill.Color = Color.FromArgb(50, Color.SteelBlue);
        chart.Line.Color = Color.SteelBlue;
        chart.Line.Width = 2;

        //chart.Legend = "X Axis: Year.\nY Axis: Average";
        chart.Legend = "X Axis: Websites \nY Axis: Validations";
        //looping through datatable and adding to chart control
        foreach (DataRow dr2 in dt.Rows)
        {
            string str = dr2["Website"].ToString();
            chart.Data.Add(new ChartPoint(dr2["Website"].ToString(), (float)System.Convert.ToSingle(dr2["Validation"])));
        }
        ConfigureColors();
        ChartControl1.Charts.Add(chart);
        ChartControl1.RedrawChart();
    }

    public System.Data.DataTable ConvertToDataTable(string path)
    {
        System.Data.DataTable dt = null;
        try
        {
            object rowIndex = 14;
            dt = new System.Data.DataTable();
            DataRow row;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(path, 0, true, 5, "", "", true,
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
            int temp = 1;
            while (((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, temp]).Value2 != null)
            {
                dt.Columns.Add(Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, temp]).Value2));
                temp++;
            }
            rowIndex = Convert.ToInt32(rowIndex) + 1;
            int columnCount = temp;
            temp = 1;
            while (((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, temp]).Value2 != null)
            {
                row = dt.NewRow();
                for (int i = 1; i < columnCount; i++)
                {
                    row[i - 1] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, i]).Value2);
                }
                dt.Rows.Add(row);
                rowIndex = Convert.ToInt32(rowIndex) + 1;
                temp = 1;
            }
            app.Workbooks.Close();
        }
        catch (Exception ex)
        {
            //lblError.Text = ex.Message;
        }
        return dt;
    }

    private void LoadExcel2()
    {
        DataTable dt = new DataTable("Chart");
        DataColumn dc = new DataColumn("Website", typeof(string));
        //DataColumn dc1 = new DataColumn("Validation", typeof(string));
        dt.Columns.Add(dc);
        //dt.Columns.Add(dc1);
        //DataRow dr = dt.NewRow();

        Excel.Application xlApp = null;
        Excel.Workbook xlWorkBook = null;
        Excel.Worksheet xlWorkSheet = null;
        Excel.Range range = null;
        //Excel.Sheets sheets;

        //string str1, str2;
        //int rCnt = 0;
        //int cCnt = 0;
        //System.Array[] myvalues;
        //string[] strArray;
        object obj;
        try
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(path1, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
            //sheets = xlWorkBook.Worksheets;
            //xlWorkSheet = (Excel.Worksheet)sheets.get_Item(1);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
            //str1 = xlWorkBook.Name;
            //str2 = xlWorkSheet.Name;

            int firstRow = 15;
            int lastRow = 1331;
            for (int i = firstRow; i <= lastRow; i++)
            {
                range = xlWorkSheet.get_Range("A" + i.ToString(), "Z" + i.ToString());
                obj = range.Cells.Value;
                System.Array myvalues = (System.Array)obj;

                //myvalues = (System.Array)range.Cells.Value;
                //strArray = ConvertToStringArray(myvalues);
                //string str = strArray[1];
                DataRow dr1 = dt.NewRow();
                string str = myvalues.GetValue(1, 2).ToString();
                dr1[0] = str;
                dt.Rows.Add(dr1);
            }


            var result = dt.AsEnumerable()
                   .GroupBy(r => r.Field<string>("Website"), StringComparer.InvariantCultureIgnoreCase)
                   .Select(r => new
                   {
                       Str = r.Key,
                       Count = r.Count()
                   });

            //LineChart chart = new LineChart();//Class instance for LineChart
            //SmoothLineChart chart = new SmoothLineChart();//Class instance for LineChart
            //AreaChart chart = new AreaChart();//Class instance for LineChart
            ColumnChart chart = new ColumnChart();
            //ScatterChart chart = new ScatterChart();
            //StackedColumnChart chart = new StackedColumnChart();

            chart.Fill.Color = Color.FromArgb(50, Color.SteelBlue);
            chart.Fill.Type = InteriorType.Hatch;
            chart.Fill.HatchStyle = System.Drawing.Drawing2D.HatchStyle.DottedDiamond;
            chart.Fill.ForeColor = Color.DarkBlue;
            chart.Line.Width = 1;
            chart.Line.Color = Color.SteelBlue;
            chart.ShowLineMarkers = true;
            chart.DataLabels.Visible = true;
            //chart.DataLabels.ShowValue = true;
            //chart.DataLabels.ShowXTitle = true;
            //chart.DataXValueField = "Websites"; ;
            //chart.DataYValueField = "Validations";
            //chart.Legend = "X Axis: Websites \nY Axis: Validations";
            //chart.ShowLegend = true;

            //looping through datatable and adding to chart control

            foreach (var item in result)
            {
                string str = item.Str;
                int i = item.Count;
                chart.Data.Add(new ChartPoint(str, (float)System.Convert.ToSingle(i)));
                //Console.WriteLine($"{item.Str} : {item.Count}");
            }

            //chart.DataLabels.ForeColor = Color.White;
            //chart.DataLabels.NumberFormat = "$0.00";
            //chart.DataLabels.Position = DataLabelPosition.Center;
            //chart.DataLabels.Background.Color = Color.SteelBlue;
            //chart.DataLabels.Border.Color = Color.Red;
            //chart.DataLabels.Separator = " ";
            //chart.DataLabels.ShowLegend = true;
            //chart.DataLabels.MaxPointsWidth = 1;
            //chart.ShowLineMarkers = true;
            //chart.LineMarker = new XLineMarker(6, Color.Red, Color.Green);
            

            //foreach (DataRow dr2 in dt.Rows)
            //{
            //    string str = dr2["Website"].ToString();
            //    chart.Data.Add(new ChartPoint(dr2["Website"].ToString(), (float)System.Convert.ToSingle(dr2["Validation"])));
            //}
            //ChartEngine engine = new ChartEngine();
            //ChartCollection charts = new ChartCollection(engine);
            //engine.Charts = charts;
            //engine.TopChartPadding = 20;
            //charts.Add(chart);


            ConfigureColors();
            ChartControl1.Charts.Clear();
            ChartControl1.Charts.Add(chart);
            ChartControl1.RedrawChart();


            xlWorkBook.Close(false, Missing.Value, Missing.Value);
            xlApp.Quit();
        }
        catch (Exception e)
        {
            string str = e.Message;
        }
        finally
        {
            // Close the Excel process
            releaseObject(range);
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
    }

    private void LoadExcel1()
    {
        //DataTable dt = new DataTable("Chart");
        //DataColumn dc = new DataColumn("Website", typeof(string));
        //DataColumn dc1 = new DataColumn("Validation", typeof(int));
        //dt.Columns.Add(dc);
        //dt.Columns.Add(dc1);
        //DataRow dr = dt.NewRow();

        Excel.Application xlApp = null;
        Excel.Workbook xlWorkBook = null;
        Excel.Worksheet xlWorkSheet = null;
        Excel.Range xlRange = null;
        //Excel.Sheets sheets;

        //string str1, str2;
        //int rCnt = 0;
        //int cCnt = 0;
        //System.Array[] myvalues;
        string[] strArray;
        object obj;
        try
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(path1, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
            //sheets = xlWorkBook.Worksheets;
            //xlWorkSheet = (Excel.Worksheet)sheets.get_Item(1);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
            //str1 = xlWorkBook.Name;
            //str2 = xlWorkSheet.Name;
            for (int i = 15; i <= 16; i++)
            {
                xlRange = xlWorkSheet.get_Range("A" + i.ToString(), "Z" + i.ToString());
                obj = xlRange.Cells.Value;
                System.Array myvalues = (System.Array)obj;
                //myvalues = (System.Array)range.Cells.Value;
                strArray = ConvertToStringArray(myvalues);
            }
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();



        }
        catch (Exception e)
        {
            string str = e.Message;
        }
        finally
        {
            // Close the Excel process
            releaseObject(xlRange);
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
    }

    private void ConfigureColors()
    {

        ChartControl1.Background.Color = Color.FromArgb(75, Color.SteelBlue);
        ChartControl1.Background.Type = InteriorType.LinearGradient;
        ChartControl1.Background.ForeColor = Color.SteelBlue;
        ChartControl1.Background.EndPoint = new Point(500, 350);

        ChartControl1.HasChartLegend = false;
        //ChartControl1.Legend.Position = LegendPosition.Bottom;
        //ChartControl1.Legend.Width = 40;

        ChartControl1.YAxisFont.ForeColor = Color.SteelBlue;
        ChartControl1.XAxisFont.ForeColor = Color.SteelBlue;
        //ChartControl1.YAxisFont.ForeColor = Color.Black;
        //ChartControl1.XAxisFont.ForeColor = Color.Black;

        ChartControl1.ChartTitle.Text = "DASHBOARD Pricing & Policy Voilation - Validations";
        ChartControl1.ChartTitle.ForeColor = Color.White;

        ChartControl1.Border.Color = Color.SteelBlue;
        ChartControl1.BorderStyle = BorderStyle.Ridge;

        ChartControl1.YCustomEnd = 200;

        ChartControl1.XTitle.Text = "Websites";
        ChartControl1.XTitle.StringFormat.Alignment = StringAlignment.Center;

        ChartControl1.YTitle.Text = "Validations";
        ChartControl1.YTitle.StringFormat.FormatFlags = StringFormatFlags.DirectionVertical;
        ChartControl1.YTitle.StringFormat.Alignment = StringAlignment.Center;


        ChartControl1.XAxisFont.StringFormat.FormatFlags = StringFormatFlags.DirectionVertical | StringFormatFlags.NoWrap | StringFormatFlags.NoClip;
        ChartControl1.XAxisFont.StringFormat.LineAlignment = StringAlignment.Center;
        ChartControl1.XAxisFont.StringFormat.Alignment = StringAlignment.Center;
        ChartControl1.XAxisFont.StringFormat.Trimming = StringTrimming.Character;
        ChartControl1.XAxisFont.ForeColor = Color.SteelBlue;
        ChartControl1.YAxisFont.ForeColor = Color.SteelBlue;

        ChartControl1.ToolTip = "Websites: No. of Validations";

        //ChartControl1.ViewStateMode = System.Web.UI.ViewStateMode.Enabled;
        //ChartControl1.XValuesInterval = 100;
        //ChartControl1.XTicksInterval = 300;


        //ChartControl1.RenderHorizontally = true;
        //ChartControl1.Width = 500;
        //ChartControl1.LeftChartPadding = 120;
        //ChartControl1.BottomChartPadding = 10;

        ChartControl1.RenderHorizontally = false;
        ChartControl1.Width = 700;
        ChartControl1.LeftChartPadding = 10;
        ChartControl1.BottomChartPadding = 120;
    }

    private void releaseObject(object obj)
    {
        try
        {
            if (null != obj)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
        catch (Exception ex)
        {
            obj = null;
            //MessageBox.Show("Unable to release the Object " + ex.ToString());
        }
        finally
        {
            GC.Collect();
            //GC.WaitForPendingFinalizers();
        }
    }

    private string[] ConvertToStringArray(System.Array values)
    {

        // create a new string array
        string[] theArray = new string[values.Length];

        // loop through the 2-D System.Array and populate the 1-D String Array
        for (int i = 1; i <= values.Length; i++)
        {
            if (values.GetValue(1, i) == null)
                theArray[i - 1] = "";
            else
                theArray[i - 1] = (string)values.GetValue(1, i).ToString();
        }

        return theArray;
    }
}