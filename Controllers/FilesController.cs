﻿using System.Data;
using System.Web;
using System.Web.Mvc;
using System.IO;
using FileHandler.Extensions;
using OfficeOpenXml;
using FileHandler.Models;
using System.Collections.Generic;
using System.Net.Http;
using System;

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
            var dt = ToDataTable(importFile);
            var rows = Duplicaterows(dt);
            if(rows.Count > 0)
            {
                return Content(@"<div class='toast align-items-center text-white bg-warning border-0 fade show' role='alert' aria-live='assertive' aria-atomic='true'>
                                      <div class='d-flex'>
                                         <div class='toast-body'>
                                         Please Check For Duplicate Records    
                                         </div>
                                         <button type='button' class='btn-close btn-close-white me-2 m-auto' data-bs-dismiss='toast' aria-label='Close'>
                                         </button>
                                      </div>
                                 </div>", "text/html");
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
        public void DeleteCustomers(string ids)
        {
            var idArray = ids.Split(',');

            foreach (var id in idArray)
            {
                CustomerCrud.DeleteCustomers(int.Parse(id));
            }
        }


        private DataTable ToDataTable(HttpPostedFileBase importFile)
        {
            string fileName = Path.GetFileName(importFile.FileName);
            string filePath = Path.Combine(Server.MapPath("~/Files"), fileName);
            importFile.SaveAs(filePath);

            DataTable dt = new DataTable();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                for (int col = 1; col <= colCount; col++)
                {
                    string columnName = worksheet.Cells[1, col].Value?.ToString();
                    if (!string.IsNullOrEmpty(columnName))
                        dt.Columns.Add(columnName);
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    DataRow dataRow = dt.NewRow();
                    for (int col = 1; col <= colCount; col++)
                    {
                        dataRow[col - 1] = worksheet.Cells[row, col].Value?.ToString();
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

                if (customerCodes.Contains(customerCode))
                {
                    duplicateCodes.Add(customerCode);
                }
                else
                {
                    customerCodes.Add(customerCode);
                }
            }

            return duplicateCodes;
        }


    }
}