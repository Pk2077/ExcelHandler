using System.Data;
using System.Web;
using System.Web.Mvc;
using System.IO;
using FileHandler.Extensions;
using OfficeOpenXml;
using System.Collections.Generic;
using System;
using System.Linq;
using OfficeOpenXml.Style;
using System.Drawing;

namespace FileHandler.Controllers
{
    public class FilesController : Controller
    {
        public ActionResult Index()
        {
            return View("Files");
        }

        [HttpGet]
        public ActionResult List()
        {
            var listModel = CustomerCrud.GetCustomers();
            return PartialView("List", listModel);
        }

        [HttpPost]
        public ActionResult FileHandler(HttpPostedFileBase importFile)
        {
            var dt = ToDataTableDynamic(importFile);
            var rows = Duplicaterows(dt);
            if(rows.Count > 0)
            {
                string duplicateRowsMessage = string.Join("<br>", rows.Select(r => r.ToString()));
                return Content($"<div class='toast align-items-center text-white bg-warning border-0 fade show' role='alert' aria-live='assertive' aria-atomic='true'><div class='d-flex'><div class='toast-body'>Please Check For Duplicate Records:<br/>{duplicateRowsMessage}</div><button type='button' class='btn-close btn-close-white me-2 m-auto' data-bs-dismiss='toast' aria-label='Close'></button></div></div>", "text/html");

            }
            else
            {
                foreach (DataRow row in dt.Rows)
                {
                    if (!ValidateRow(row))
                    {
                        CustomerCrud.InsertCustomers(row);
                    }
                }
                return List();
            }

        }
        [HttpPost]
        public ActionResult ExportToExcel(List<List<Dictionary<string, object>>> tableData)
        {
            // Create Excel package and worksheet
            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");

            // Set headers
            string[] headers = { "Id", "Name", "Code", "Address 1", "Address 2", "City", "State", "Pin", "Mobile No" };
            for (int col = 0; col < headers.Length; col++)
            {
                workSheet.Cells[1, col + 1].Value = headers[col];
            }

            // Populate data and apply styles
            for (int row = 0; row < tableData.Count; row++)
            {
                for (int col = 0; col < tableData[row].Count; col++)
                {
                    var cellData = tableData[row][col];
                    if (cellData != null && cellData.ContainsKey("value"))
                    {
                        workSheet.Cells[row + 2, col + 1].Value = cellData["value"];

                        if (cellData.ContainsKey("backgroundColor") && cellData["backgroundColor"] != null)
                        {
                            string colorValue = cellData["backgroundColor"].ToString();
                            Color color = ParseColor(colorValue);
                            if (color != null)
                            {
                                workSheet.Cells[row + 2, col + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[row + 2, col + 1].Style.Fill.BackgroundColor.SetColor(color);
                            }
                        }
                    }
                }
            }

            // Prepare response
            byte[] fileContents = excel.GetAsByteArray();
            string fileName = "table_data.xlsx";

            // Return Excel file as downloadable response
            return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        // Helper method to parse color string to Color object
        private Color ParseColor(string colorValue)
        {
            try
            {
                // Check if colorValue is a named color (e.g., "Red", "Blue", etc.)
                Color namedColor = Color.FromName(colorValue);
                if (namedColor.IsKnownColor)
                {
                    return namedColor;
                }

                // Try parsing as a hexadecimal color (e.g., "#RRGGBB" format)
                if (colorValue.StartsWith("#") && colorValue.Length == 7)
                {
                    return ColorTranslator.FromHtml(colorValue);
                }

                return Color.FromName("white"); // Invalid color format
            }
            catch
            {
                return Color.FromName("white"); // Error parsing color
            }
        }

        [HttpPost]
        public void DeleteCustomers(string ids)
        {
            var idArray = ids.Split(',');

            foreach (var id in idArray)
            {
                CustomerCrud.DeleteCustomers(int.Parse(id));
            }
        }

        private DataTable ToDataTableDynamic(HttpPostedFileBase importFile)
        {
            string fileName = Path.GetFileName(importFile.FileName);
            string filePath = Path.Combine(Server.MapPath("~/Files"), fileName);
            importFile.SaveAs(filePath);

            DataTable dt = new DataTable();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.End.Row;
                int colCount = worksheet.Dimension.Columns;

                int startRow = 1;
                for (int row = 1; row <= rowCount; row++)
                {
                    bool isEmptyRow = true;
                    for (int col = 1; col <= colCount; col++)
                    {
                        if (worksheet.Cells[row, col].Value != null)
                        {
                            isEmptyRow = false;
                            break;
                        }
                    }
                    if (!isEmptyRow)
                    {
                        startRow = row;
                        break;
                    }
                }

                int startCol = 1;
                for (int col = 1; col <= colCount; col++)
                {
                    bool isEmptyColumn = true;
                    for (int row = startRow; row <= rowCount; row++)
                    {
                        if (worksheet.Cells[row, col].Value != null)
                        {
                            isEmptyColumn = false;
                            break;
                        }
                    }
                    if (!isEmptyColumn)
                    {
                        startCol = col;
                        break;
                    }
                }

                for (int col = startCol; col <= colCount; col++)
                {
                    string columnName = worksheet.Cells[startRow, col].Value?.ToString();
                    if (!string.IsNullOrEmpty(columnName))
                        dt.Columns.Add(columnName);
                }

                for (int row = startRow + 1; row <= rowCount; row++)
                {
                    DataRow dataRow = dt.NewRow();
                    for (int col = startCol; col <= colCount; col++)
                    {
                        dataRow[col - startCol] = worksheet.Cells[row, col].Value?.ToString();
                    }
                    dt.Rows.Add(dataRow);
                }
            }

            return dt;
        }

        private bool ValidateRow(DataRow row)
        {
           var customersInDb = CustomerCrud.GetCustomersByCode(row["Customer Code"].ToString());
            if(customersInDb != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private HashSet<string> Duplicaterows(DataTable dt)
        {
            HashSet<string> customerCodes = new HashSet<string>();
            HashSet<string> duplicateCodes = new HashSet<string>();

            foreach (DataRow row in dt.Rows)
            {
                string customerCode = row["Customer Code"].ToString();
                if(!string.IsNullOrEmpty(customerCode))
                {
                    if (customerCodes.Contains(customerCode))
                        duplicateCodes.Add(customerCode);
                    else
                        customerCodes.Add(customerCode);
                }
                
            }
            return duplicateCodes;
        }


    }
}