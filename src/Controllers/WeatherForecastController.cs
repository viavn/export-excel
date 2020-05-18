using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using Microsoft.AspNetCore.Mvc;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.UserModel.Charts;

namespace ExcelExportExample.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        const int NUM_OF_ROWS = 3;
        const int NUM_OF_COLUMNS = 10;

        [HttpGet]
        public IEnumerable<Praca> Get()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            var pracas = GetPracas();

            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("linechart");

            // Create a row and put some cells in it. Rows are 0 based.
            IRow row;
            ICell cell;

            IDataFormat format = wb.CreateDataFormat();
            int rowIndex = 0;

            foreach (var praca in pracas)
            {
                row = sheet.CreateRow(rowIndex);
                rowIndex++;

                int cellIndex = 0;
                cell = row.CreateCell(cellIndex++);
                cell.SetCellValue($"Row{rowIndex}");

                foreach (var passagem in praca.Passagens)
                {
                    cell = row.CreateCell(cellIndex++);
                    SetValueAndFormat(wb, cell, passagem.Abono, format.GetFormat("0.00%"));

                    cell = row.CreateCell(cellIndex++);
                    SetValueAndFormat(wb, cell, passagem.Isento, format.GetFormat("0.00%"));

                    cell = row.CreateCell(cellIndex);
                    SetValueAndFormat(wb, cell, passagem.Violacao, format.GetFormat("0.00%"));
                }
            }

            CreateChart(sheet);

            using (FileStream fs = System.IO.File.Create("E:\\test.xlsx"))
            {
                wb.Write(fs);
            }

            //using (FileStream fs = System.IO.File.Create("E:\\test.xlsx"))
            //{
            //    var wb = CreateBarchart();
            //    wb.Write(fs);
            //}

            return pracas;
        }

        static void SetValueAndFormat(IWorkbook workbook, ICell cell, double value, short formatId)
        {
            cell.SetCellValue(value);
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.DataFormat = formatId;
            cell.CellStyle = cellStyle;
        }

        static void CreateChart(ISheet sheet)
        {
            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFClientAnchor anchor = (XSSFClientAnchor)drawing.CreateAnchor(0, 0, 0, 0, 3, 0, 10, 15);

            IChart chart = drawing.CreateChart(anchor);
            IChartLegend legend = chart.GetOrCreateLegend();
            legend.Position = LegendPosition.Bottom;

            IBarChartData<string, double> data = chart.ChartDataFactory.CreateBarChartData<string, double>();

            IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            bottomAxis.MajorTickMark = AxisTickMark.None;
            IValueAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
            leftAxis.Crosses = AxisCrosses.AutoZero;
            leftAxis.SetCrossBetween(AxisCrossBetween.Between);

            IChartDataSource<string> categories = DataSources.FromStringCellRange(sheet, new CellRangeAddress(0, GetPracas().Count() - 1, 0, 0));

            // Abono
            IChartDataSource<double> xAbono = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(0, GetPracas().Count() - 1, 1, 1));            
            // Isento
            IChartDataSource<double> xIsento = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(0, GetPracas().Count() - 1, 2, 2));            
            // Violação
            IChartDataSource<double> xViolacao = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(0, GetPracas().Count() - 1, 3, 3));

            var s1 = data.AddSeries(categories, xAbono);
            s1.SetTitle("Abono");
            
            var s2 = data.AddSeries(categories, xIsento);
            s2.SetTitle("Isento");
            
            var s3 = data.AddSeries(categories, xViolacao);
            s3.SetTitle("Violação");

            chart.Plot(data, bottomAxis, leftAxis);
        }

        private XSSFWorkbook CreateBarchart()
        {

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("WorkForceAnalytics");

            XSSFCellStyle styleHeader = (XSSFCellStyle)workbook.CreateCellStyle();
            styleHeader.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            //styleHeader.SetFont(getNewXSSFFont(workbook, styleHeader));

            XSSFCellStyle style = (XSSFCellStyle)workbook.CreateCellStyle();
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;

            XSSFRow row1 = (XSSFRow)sheet.CreateRow(0);
            row1.CreateCell(0).SetCellValue("");

            List<string> lstdatase = new List<string>(4)
            {
                "A1",
                "B1",
                "C1",
                "D1",
                "E1"
            };

            for (int i = 1; i < 5; i++)
            {
                row1.CreateCell(i).SetCellValue(lstdatase[i - 1].ToString());
                row1.GetCell(i).CellStyle = styleHeader;
            }

            int rowvalue = 1;
            List<string> lstdata = new List<string>(8)
            {
                "A",
                "B",
                "C",
                "D",
                "E",
                "F",
                "G",
                "H"
            };

            int d = 10;

            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFClientAnchor anchor = (XSSFClientAnchor)drawing.CreateAnchor(0, 0, 0, 0, 6, 1, 15, 18);

            IChart chart = drawing.CreateChart(anchor);
            IChartLegend legend = chart.GetOrCreateLegend();
            legend.Position = (LegendPosition.Bottom);

            IBarChartData<string, double> data = chart.ChartDataFactory.CreateBarChartData<string, double>();

            IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            IValueAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
            leftAxis.Crosses = AxisCrosses.AutoZero;
            leftAxis.SetCrossBetween(AxisCrossBetween.Between);

            IChartDataSource<string> xs = DataSources.FromStringCellRange(sheet, new NPOI.SS.Util.CellRangeAddress(0, 0, 1, 5 - 1));

            for (int ii = 0; ii < 8; ii++)
            {
                XSSFRow rownew = (XSSFRow)sheet.CreateRow(rowvalue);
                rownew.CreateCell(0).SetCellValue(lstdata[ii].ToString());

                for (int i = 1; i < 5; i++)
                {
                    rownew.CreateCell(i).SetCellValue(d * 0.1);
                    d++;
                    rownew.GetCell(i).CellStyle = style;
                }
                rowvalue++;


                IChartDataSource<double> ys = DataSources.FromNumericCellRange(sheet, new NPOI.SS.Util.CellRangeAddress(ii + 1, ii + 1, 1, 5 - 1));
                data.AddSeries(xs, ys).SetTitle(lstdata[ii].ToString());
            }

            chart.Plot(data, bottomAxis, leftAxis);

            sheet.ForceFormulaRecalculation = true;
            return workbook;
        }

        static IEnumerable<Praca> GetPracas()
        {
            var pracas = new List<Praca>();
            for (int i = 1; i < 15; i++)
            {
                var passagens = new List<Passagem>
                {
                    new Passagem { Abono =  (i * 0.01), Isento = (i * 2 * 0.01), Violacao = (i * 3 * 0.01) }
                };

                pracas.Add(new Praca { Nome = $"Praça {i}", Passagens = passagens });
            }

            return pracas;
        }
    }
}
