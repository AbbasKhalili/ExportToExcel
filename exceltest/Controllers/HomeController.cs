using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using NPOI;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;
using Syncfusion.Licensing;
using ExcelHorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment;
using ExcelVerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace exceltest.Controllers
{
    public class HomeController : Controller
    {
        public async Task<IActionResult> Index()
        {
            var data = DataInfo.Getdata();

            #region Syncfusion

            //=============================================Syncfusion

            //using (ExcelEngine excelEngine = new ExcelEngine())
            //{
            //	IApplication application = excelEngine.Excel;

            //	application.DefaultVersion = ExcelVersion.Excel2016;

            //	//Create a workbook
            //	IWorkbook workbook = application.Workbooks.Create(1);
            //	IWorksheet worksheet = workbook.Worksheets[0];

            //	//Adding a picture
            //	//FileStream imageStream = new FileStream("AdventureCycles-Logo.png", FileMode.Open, FileAccess.Read);
            //	//IPictureShape shape = worksheet.Pictures.AddPicture(1, 1, imageStream);

            //	//Disable gridlines in the worksheet
            //	worksheet.IsGridLinesVisible = false;

            //	//Enter values to the cells from A3 to A5
            //	worksheet.Range["A3"].Text = "46036 Michigan Ave";
            //	worksheet.Range["A4"].Text = "Canton, USA";
            //	worksheet.Range["A5"].Text = "Phone: +1 231-231-2310";

            //	//Make the text bold
            //	worksheet.Range["A3:A5"].CellStyle.Font.Bold = true;

            //	//Merge cells
            //	worksheet.Range["D1:E1"].Merge();

            //	//Enter text to the cell D1 and apply formatting.
            //	worksheet.Range["D1"].Text = "INVOICE";
            //	worksheet.Range["D1"].CellStyle.Font.Bold = true;
            //	worksheet.Range["D1"].CellStyle.Font.RGBColor = Color.FromArgb(42, 118, 189);
            //	worksheet.Range["D1"].CellStyle.Font.Size = 35;

            //	//Apply alignment in the cell D1
            //	worksheet.Range["D1"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;
            //	worksheet.Range["D1"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;

            //	//Enter values to the cells from D5 to E8
            //	worksheet.Range["D5"].Text = "INVOICE#";
            //	worksheet.Range["E5"].Text = "DATE";
            //	worksheet.Range["D6"].Number = 1028;
            //	worksheet.Range["E6"].Value = "12/31/2018";
            //	worksheet.Range["D7"].Text = "CUSTOMER ID";
            //	worksheet.Range["E7"].Text = "TERMS";
            //	worksheet.Range["D8"].Number = 564;
            //	worksheet.Range["E8"].Text = "Due Upon Receipt";

            //	//Apply RGB backcolor to the cells from D5 to E8
            //	worksheet.Range["D5:E5"].CellStyle.Color = Color.FromArgb(42, 118, 189);
            //	worksheet.Range["D7:E7"].CellStyle.Color = Color.FromArgb(42, 118, 189);

            //	//Apply known colors to the text in cells D5 to E8
            //	worksheet.Range["D5:E5"].CellStyle.Font.Color = ExcelKnownColors.White;
            //	worksheet.Range["D7:E7"].CellStyle.Font.Color = ExcelKnownColors.White;

            //	//Make the text as bold from D5 to E8
            //	worksheet.Range["D5:E8"].CellStyle.Font.Bold = true;

            //	//Apply alignment to the cells from D5 to E8
            //	worksheet.Range["D5:E8"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            //	worksheet.Range["D5:E5"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
            //	worksheet.Range["D7:E7"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
            //	worksheet.Range["D6:E6"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;

            //	//Enter value and applying formatting in the cell A7
            //	worksheet.Range["A7"].Text = "  BILL TO";
            //	worksheet.Range["A7"].CellStyle.Color = Color.FromArgb(42, 118, 189);
            //	worksheet.Range["A7"].CellStyle.Font.Bold = true;
            //	worksheet.Range["A7"].CellStyle.Font.Color = ExcelKnownColors.White;

            //	//Apply alignment
            //	worksheet.Range["A7"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
            //	worksheet.Range["A7"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

            //	//Enter values in the cells A8 to A12
            //	worksheet.Range["A8"].Text = "Steyn";
            //	worksheet.Range["A9"].Text = "Great Lakes Food Market";
            //	worksheet.Range["A10"].Text = "20 Whitehall Rd";
            //	worksheet.Range["A11"].Text = "North Muskegon,USA";
            //	worksheet.Range["A12"].Text = "+1 231-654-0000";

            //	//Create a Hyperlink for e-mail in the cell A13
            //	IHyperLink hyperlink = worksheet.HyperLinks.Add(worksheet.Range["A13"]);
            //	hyperlink.Type = ExcelHyperLinkType.Url;
            //	hyperlink.Address = "Steyn@greatlakes.com";
            //	hyperlink.ScreenTip = "Send Mail";

            //	//Merge column A and B from row 15 to 22
            //	worksheet.Range["A15:B15"].Merge();
            //	worksheet.Range["A16:B16"].Merge();
            //	worksheet.Range["A17:B17"].Merge();
            //	worksheet.Range["A18:B18"].Merge();
            //	worksheet.Range["A19:B19"].Merge();
            //	worksheet.Range["A20:B20"].Merge();
            //	worksheet.Range["A21:B21"].Merge();
            //	worksheet.Range["A22:B22"].Merge();

            //	//Enter details of products and prices
            //	worksheet.Range["A15"].Text = "  DESCRIPTION";
            //	worksheet.Range["C15"].Text = "QTY";
            //	worksheet.Range["D15"].Text = "UNIT PRICE";
            //	worksheet.Range["E15"].Text = "AMOUNT";
            //	worksheet.Range["A16"].Text = "Cabrales Cheese";
            //	worksheet.Range["A17"].Text = "Chocos";
            //	worksheet.Range["A18"].Text = "Pasta";
            //	worksheet.Range["A19"].Text = "Cereals";
            //	worksheet.Range["A20"].Text = "Ice Cream";
            //	worksheet.Range["C16"].Number = 3;
            //	worksheet.Range["C17"].Number = 2;
            //	worksheet.Range["C18"].Number = 1;
            //	worksheet.Range["C19"].Number = 4;
            //	worksheet.Range["C20"].Number = 3;
            //	worksheet.Range["D16"].Number = 21;
            //	worksheet.Range["D17"].Number = 54;
            //	worksheet.Range["D18"].Number = 10;
            //	worksheet.Range["D19"].Number = 20;
            //	worksheet.Range["D20"].Number = 30;
            //	worksheet.Range["D23"].Text = "Total";

            //	//Apply number format
            //	worksheet.Range["D16:E22"].NumberFormat = "$.00";
            //	worksheet.Range["E23"].NumberFormat = "$.00";

            //	//Apply incremental formula for column Amount by multiplying Qty and UnitPrice
            //	application.EnableIncrementalFormula = true;
            //	worksheet.Range["E16:E20"].Formula = "=C16*D16";

            //	//Formula for Sum the total
            //	worksheet.Range["E23"].Formula = "=SUM(E16:E22)";

            //	//Apply borders
            //	worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
            //	worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
            //	worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Grey_25_percent;
            //	worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Grey_25_percent;
            //	worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
            //	worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
            //	worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Black;
            //	worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Black;

            //	//Apply font setting for cells with product details
            //	worksheet.Range["A3:E23"].CellStyle.Font.FontName = "Arial";
            //	worksheet.Range["A3:E23"].CellStyle.Font.Size = 10;
            //	worksheet.Range["A15:E15"].CellStyle.Font.Color = ExcelKnownColors.White;
            //	worksheet.Range["A15:E15"].CellStyle.Font.Bold = true;
            //	worksheet.Range["D23:E23"].CellStyle.Font.Bold = true;

            //	//Apply cell color
            //	worksheet.Range["A15:E15"].CellStyle.Color = Color.FromArgb(42, 118, 189);

            //	//Apply alignment to cells with product details
            //	worksheet.Range["A15"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
            //	worksheet.Range["C15:C22"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            //	worksheet.Range["D15:E15"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;

            //	//Apply row height and column width to look good
            //	worksheet.Range["A1"].ColumnWidth = 36;
            //	worksheet.Range["B1"].ColumnWidth = 11;
            //	worksheet.Range["C1"].ColumnWidth = 8;
            //	worksheet.Range["D1:E1"].ColumnWidth = 18;
            //	worksheet.Range["A1"].RowHeight = 47;
            //	worksheet.Range["A2"].RowHeight = 15;
            //	worksheet.Range["A3:A4"].RowHeight = 15;
            //	worksheet.Range["A5"].RowHeight = 18;
            //	worksheet.Range["A6"].RowHeight = 29;
            //	worksheet.Range["A7"].RowHeight = 18;
            //	worksheet.Range["A8"].RowHeight = 15;
            //	worksheet.Range["A9:A14"].RowHeight = 15;
            //	worksheet.Range["A15:A23"].RowHeight = 18;


            //	//Saving the Excel to the MemoryStream 
            //	MemoryStream stream = new MemoryStream();

            //	workbook.SaveAs(stream);

            //	//Set the position as '0'.
            //	stream.Position = 0;

            //	//Download the Excel file in the browser
            //             FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel")
            //             {
            //                 FileDownloadName = "Output.xlsx"
            //             };


            //             return fileStreamResult;
            //}

            #endregion


            #region PEPlus

            //============================================= PEPlus

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("EquipmentRevenue");

                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Font.Size = 10;

                worksheet.Cells["A1:L1"].Merge = true;
                worksheet.Cells["A1"].Value = "Equipment Revenue - MBS EQUIPMENT CO.";
                worksheet.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                worksheet.Cells["A1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                worksheet.Cells["A1"].Style.Font.Bold = true;
                worksheet.Cells["A1"].Style.Font.UnderLine = true;
                worksheet.Cells["A1"].Style.Font.Size = 12;

                worksheet.Cells["A2:L2"].Merge = true;
                worksheet.Cells["A2"].Value = "Date Range: 3/2/2020 - 3/2/2020\nInclude Loss & Damage: Yes";
                worksheet.Cells["A2"].Style.Font.Bold = true;
                worksheet.Cells["A2"].Style.WrapText = true;
                worksheet.Row(2).Height = 30;



                //worksheet.Cells["A3"].LoadFromCollection(data, true, TableStyles.Light1);

                var props = data.GetType().GetGenericArguments().Single().GetProperties();
                for (var i = 0; i < props.Length; i++)
                {
                    var value = props[i].Name;
                    var displayName = props[i].CustomAttributes
                        .FirstOrDefault(a => a.AttributeType == typeof(DisplayNameAttribute));
                    if (displayName != null)
                        value = displayName.ConstructorArguments.FirstOrDefault().Value.ToString();


                    worksheet.Cells[3, i + 1].Style.WrapText = true;
                    worksheet.Cells[3, i + 1].Value = value?.Replace("&", "\n");

                    var colSize = 15;
                    var size = props[i].CustomAttributes
                        .FirstOrDefault(a => a.AttributeType == typeof(StringLengthAttribute));
                    if (size != null)
                        colSize = (int)size.ConstructorArguments.FirstOrDefault().Value;

                    worksheet.Column(i + 1).Width = colSize;
                }

                worksheet.Row(3).Height = 37;
                worksheet.Row(3).Style.Font.Bold = true;
                worksheet.Row(3).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Row(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[3, 1, 3, 12].AutoFilter = true;


                var startRow = 4;
                var row = 4;
                foreach (var itm in data)
                {
                    worksheet.Cells[row, 1].Value = itm.Today;
                    worksheet.Cells[row, 1].Style.Numberformat.Format = "mm-dd-yyyy";

                    worksheet.Cells[row, 2].Value = itm.Equipment;
                    worksheet.Cells[row, 3].Value = itm.Department;
                    worksheet.Cells[row, 4].Value = itm.Category;
                    worksheet.Cells[row, 5].Value = itm.EquipDesc;
                    worksheet.Cells[row, 6].Value = itm.OrderNo;
                    worksheet.Cells[row, 7].Value = itm.OrderDesc;
                    worksheet.Cells[row, 8].Value = itm.Vendor;
                    worksheet.Cells[row, 9].Value = itm.DaysRented;
                    worksheet.Cells[row, 10].Value = itm.QtyRented;
                    worksheet.Cells[row, 11].Value = itm.QtyBilled;
                    worksheet.Cells[row, 12].Value = itm.DaysBilled;
                    row++;
                }
                worksheet.Cells[row, 6].Formula = $"=SUM(F{startRow}:F{row - 1})";
                worksheet.Cells[row, 9].Formula = $"=SUM(I{startRow}:I{row - 1})";
                worksheet.Cells[row, 10].Formula = $"=SUM(J{startRow}:J{row - 1})";
                worksheet.Cells[row, 11].Formula = $"=SUM(K{startRow}:K{row - 1})";
                worksheet.Cells[row, 12].Formula = $"=SUM(L{startRow}:L{row - 1})";

                worksheet.Cells[row, 1, row, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, 1, row, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                worksheet.Cells[row, 1, row, 12].Style.Font.Color.SetColor(System.Drawing.Color.White);
                worksheet.Cells[row, 1, row, 12].Style.Font.Bold = true;

                //worksheet.Cells.AutoFitColumns(13);

                worksheet.View.FreezePanes(4, 1);

                //worksheet.Cells["A7:E7"].Merge = true;
                //worksheet.Cells["A7"].Value = "Location: MBS3RDRAIL \n Currency: USD";
                //worksheet.Cells["A7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //worksheet.Cells["A7"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.BlueViolet);
                //worksheet.Cells["A7"].Style.Font.Color.SetColor(System.Drawing.Color.Lime);
                //worksheet.Cells["A7"].Style.Font.Bold = true;


                ////Add a formula for the value-column
                //worksheet.Cells["E2:E4"].Formula = "C2*D2";

                ////Ok now format the values;
                //using (var range = worksheet.Cells[1, 1, 1, 5])
                //{
                //	range.Style.Font.Bold = true;
                //	range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                //	range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
                //	range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                //}

                //worksheet.Cells["A5:E5"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                //worksheet.Cells["A5:E5"].Style.Font.Bold = true;

                //worksheet.Cells[5, 3, 5, 5].Formula = string.Format("SUBTOTAL(9,{0})", new ExcelAddress(2, 3, 4, 3).Address);
                //worksheet.Cells["C2:C5"].Style.Numberformat.Format = "#,##0";
                //worksheet.Cells["D2:E5"].Style.Numberformat.Format = "#,##0.00";

                ////Create an autofilter for the range
                //worksheet.Cells["A1:E4"].AutoFilter = true;

                //worksheet.Cells["A2:A4"].Style.Numberformat.Format = "@";   //Format as text

                //There is actually no need to calculate, Excel will do it for you, but in some cases it might be useful. 
                //For example if you link to this workbook from another workbook or you will open the workbook in a program that hasn't a calculation engine or 
                //you want to use the result of a formula in your program.
                worksheet.Calculate();

                //worksheet.Cells.AutoFitColumns(0);  //Autofit columns for all cells

                // Lets set the header text 
                //worksheet.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\" Inventory";
                //// Add the page number to the footer plus the total number of pages
                //worksheet.HeaderFooter.OddFooter.RightAlignedText =
                //	string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                //// Add the sheet name to the footer
                //worksheet.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                //// Add the file path to the footer
                //worksheet.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FilePath + ExcelHeaderFooter.FileName;

                //worksheet.PrinterSettings.RepeatRows = worksheet.Cells["1:2"];
                //worksheet.PrinterSettings.RepeatColumns = worksheet.Cells["A:G"];

                // Change the sheet view to show it in page layout mode
                //worksheet.View.PageLayoutView = true;

                //// Set some document properties
                //package.Workbook.Properties.Title = "Invertory";
                //package.Workbook.Properties.Author = "Jan Källman";
                //package.Workbook.Properties.Comments = "This sample demonstrates how to create an Excel workbook using EPPlus";

                //// Set some extended property values
                //package.Workbook.Properties.Company = "EPPlus Software AB";

                //// Set some custom property values
                //package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Jan Källman");
                //package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");

                ////var xlFile = new FileInfo($"{DateTime.Now:yyyymmddhhmmss}.xlsx");

                ////// Save our new workbook in the output directory and we are done!
                ////package.SaveAs(xlFile);

                var filename = $"{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                var stream = new MemoryStream(package.GetAsByteArray());
                //package.SaveAs(stream);
                stream.Position = 0;
                var fileStreamResult = new FileStreamResult(stream,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    FileDownloadName = filename
                };
                return fileStreamResult;



                //return View();
            }

            #endregion


            #region NPOI

            //=============================================NPOI
            //var newFile = @"newbook.core.xlsx";

            //         using (var fs = new FileStream(newFile, FileMode.Create, FileAccess.Write))
            //         {

            //             var workbook = new XSSFWorkbook();

            //             ISheet sheet1 = workbook.CreateSheet("Sheet1");

            //             sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 10));
            //             var rowIndex = 0;
            //             IRow row = sheet1.CreateRow(rowIndex);
            //             row.Height = 30 * 80;
            //             row.CreateCell(0).SetCellValue("this is content");
            //             sheet1.AutoSizeColumn(0);
            //             rowIndex++;

            //             var sheet2 = workbook.CreateSheet("Sheet2");
            //             var style1 = workbook.CreateCellStyle();
            //             style1.FillForegroundColor = HSSFColor.Blue.Index2;
            //             style1.FillPattern = FillPattern.SolidForeground;

            //             var style2 = workbook.CreateCellStyle();
            //             style2.FillForegroundColor = HSSFColor.Yellow.Index2;
            //             style2.FillPattern = FillPattern.SolidForeground;

            //             var cell2 = sheet2.CreateRow(0).CreateCell(0);
            //             cell2.CellStyle = style1;
            //             cell2.SetCellValue(0);

            //             cell2 = sheet2.CreateRow(1).CreateCell(0);
            //             cell2.CellStyle = style2;
            //             cell2.SetCellValue(1);

            //             workbook.Write(fs);
            //             fs.Position = 0;
            //             var filename = $"{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            //	var fileStreamResult = new FileStreamResult(fs,
            //                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            //             {
            //                 FileDownloadName = filename
            //             };
            //             return fileStreamResult;
            //}


            //var workbook = new XSSFWorkbook();
            //var excelSheet = workbook.CreateSheet("Demo");

            //excelSheet.CreateFreezePane(2, 2);
            //excelSheet.SetColumnWidth(2, 1500);
            //excelSheet.AddMergedRegion(new CellRangeAddress(17, 18, 0, 1));

            //excelSheet.SetAutoFilter(new CellRangeAddress(1, 1, 0, 2));

            //var row = excelSheet.CreateRow(0);
            //row.Height = 500;
            //var t = row.CreateCell(0);
            //t.SetCellValue("List");

            //var f = workbook.CreateFont();
            //f.IsBold = true;
            //f.Underline = FontUnderlineType.Single;

            //t.CellStyle.SetFont(f);



            //row = excelSheet.CreateRow(1);
            //var cel = row.CreateCell(0);
            //cel.SetCellValue("ID");

            //row.CreateCell(1).SetCellValue("Name");
            //row.CreateCell(2).SetCellValue("Age");

            //row = excelSheet.CreateRow(2);
            //row.CreateCell(0).SetCellValue(1);
            //row.CreateCell(1).SetCellValue("Kane Williamson");
            //row.CreateCell(2).SetCellValue(29);

            //row = excelSheet.CreateRow(3);
            //row.CreateCell(0).SetCellValue(2);
            //row.CreateCell(1).SetCellValue("Martin Guptil");
            //row.CreateCell(2).SetCellValue(33);

            //row = excelSheet.CreateRow(4);
            //row.CreateCell(0).SetCellValue(3);
            //row.CreateCell(1).SetCellValue("Colin Munro");
            //row.CreateCell(2).SetCellValue(23);


            //POIXMLProperties props = workbook.GetProperties();
            //props.CoreProperties.Creator = "NPOI 2.5.1";
            //props.CoreProperties.Created = DateTime.Now;
            //if (!props.CustomProperties.Contains("NPOI Team"))
            //    props.CustomProperties.AddProperty("NPOI Team", "Hello World!");




            //ICellStyle rowstyle = workbook.CreateCellStyle();
            //rowstyle.FillForegroundColor = IndexedColors.Red.Index;
            //rowstyle.FillPattern = FillPattern.SolidForeground;

            //ICellStyle c1Style = workbook.CreateCellStyle();
            //c1Style.FillForegroundColor = IndexedColors.Yellow.Index;
            //c1Style.FillPattern = FillPattern.SolidForeground;

            //IRow r1 = excelSheet.CreateRow(5);
            //IRow r2 = excelSheet.CreateRow(6);
            //r1.RowStyle = rowstyle;
            //r2.RowStyle = rowstyle;

            //ICell c1 = r2.CreateCell(7);
            //c1.CellStyle = c1Style;
            //c1.SetCellValue("Test");

            //ICell c4 = r2.CreateCell(9);
            //c4.CellStyle = c1Style;


            ////font style1: underlined, italic, red color, fontsize=20
            //var font1 = workbook.CreateFont();
            //font1.Color = IndexedColors.Red.Index;
            //font1.IsItalic = true;
            //font1.Underline = FontUnderlineType.Double;
            //font1.FontHeightInPoints = 20;

            ////bind font with style 1
            //ICellStyle style1 = workbook.CreateCellStyle();
            //style1.SetFont(font1);

            ////font style2: strikeout line, green color, fontsize=15, fontname='宋体'
            //var font2 = workbook.CreateFont();
            //font2.Color = IndexedColors.OliveGreen.Index;
            //font2.IsStrikeout = true;
            //font2.FontHeightInPoints = 22;
            //font2.FontName = "Tahoma";

            ////bind font with style 2
            //ICellStyle style2 = workbook.CreateCellStyle();
            //style2.SetFont(font2);

            ////apply font styles
            //ICell cell1 = excelSheet.CreateRow(10).CreateCell(1);
            //cell1.SetCellValue("Hello World!");
            //cell1.CellStyle = style1;
            //ICell cell2 = excelSheet.CreateRow(11).CreateCell(1);
            //cell2.SetCellValue("wow");
            //cell2.CellStyle = style2;

            //////cell with rich text 
            //ICell cell3 = excelSheet.CreateRow(12).CreateCell(1);
            //XSSFRichTextString richtext = new XSSFRichTextString("Microsoft OfficeTM");

            ////apply font to "Microsoft Office"
            //var font4 = workbook.CreateFont();
            //font4.FontHeightInPoints = 12;
            //richtext.ApplyFont(0, 16, font4);
            ////apply font to "TM"
            //var font3 = workbook.CreateFont();
            //font3.TypeOffset = FontSuperScript.Super;
            //font3.IsItalic = true;
            //font3.Color = IndexedColors.Blue.Index;
            //font3.FontHeightInPoints = 8;
            //richtext.ApplyFont(16, 18, font3);

            //cell3.SetCellValue(richtext);


            //var filename = $"{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            //var stream = new MemoryStream();
            //workbook.Write(stream);
            //stream.Position = 0;
            //var fileStreamResult = new FileStreamResult(stream,
            //    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            //{
            //    FileDownloadName = filename
            //};
            //return fileStreamResult;


            #endregion
        }

    }
}




