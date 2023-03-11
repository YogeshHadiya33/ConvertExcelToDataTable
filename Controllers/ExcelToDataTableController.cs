using ExcelToDataTable.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Data;

namespace ExcelToDataTable.Controllers
{
    public class ExcelToDataTableController : Controller
    {
        public IActionResult Index()
        {
            FileUploadModel model = new FileUploadModel();
            return View(model);
        }

        [HttpPost]
        public IActionResult Convert(FileUploadModel model)
        {
            DataTable table = new DataTable();
            try
            {
                if (model.File != null)
                {
                    //if you want to read data from a excel file use this 
                    //using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                    using (var stream = model.File.OpenReadStream())
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                        ExcelPackage package = new ExcelPackage();
                        package.Load(stream);
                        if (package.Workbook.Worksheets.Count > 0)
                        {
                            using (ExcelWorksheet workSheet = package.Workbook.Worksheets.First())
                            {
                                int noOfCol = workSheet.Dimension.End.Column;
                                int noOfRow = workSheet.Dimension.End.Row;
                                int rowIndex = 1;

                                for (int c = 1; c <= noOfCol; c++)
                                {
                                    table.Columns.Add(workSheet.Cells[rowIndex, c].Text);
                                }
                                rowIndex = 2;
                                for (int r = rowIndex; r <= noOfRow; r++)
                                {
                                    DataRow dr = table.NewRow();
                                    for (int c = 1; c <= noOfCol; c++)
                                    {
                                        dr[c - 1] = workSheet.Cells[r, c].Value;
                                    }
                                    table.Rows.Add(dr);
                                }

                                ViewBag.SuccessMessage = "Excel Successfully Converted to Data Table";
                            }
                        }
                        else
                            ViewBag.ErrorMessage = "No Work Sheet available in Excel File";

                    }
                }

            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = ex.Message;
            }
            return View("Index");
        }

    }
}
