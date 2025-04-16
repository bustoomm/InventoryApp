using Microsoft.AspNetCore.Mvc;
using InventoryApp.Data;
using InventoryApp.Models;
using Dapper;
using System.IO;
using ClosedXML.Excel;
using DinkToPdf.Contracts;
using DinkToPdf;
using InventoryApp.Helpers;

namespace InventoryApp.Controllers
{
    public class InventoryController : Controller
    {
        private readonly Db _db;
        private readonly IConverter _converter;

        public InventoryController(Db db, IConverter converter)
        {
            _db = db;
            _converter = converter;
        }

        public IActionResult Create()
        {
            if (HttpContext.Session.GetString("role") != "admin")
                return Unauthorized();

            return View();
        }

        [HttpPost]
        public IActionResult Create(InventoryItem item)
        {
            using (var conn = _db.GetConnection())
            {
                conn.Open();
                string query = @"INSERT INTO InventoryItems 
                                 (ItemCode, ItemName, Quantity, Unit, Location) 
                                 VALUES (@ItemCode, @ItemName, @Quantity, @Unit, @Location)";

                conn.Execute(query, item);
            }

            return RedirectToAction("Index");
        }

        public IActionResult Index(string search)
        {
            using (var conn = _db.GetConnection())
            {
                conn.Open();
                string query = "SELECT * FROM InventoryItems";

                if (!string.IsNullOrEmpty(search))
                {
                    query += " WHERE ItemName LIKE @search OR ItemCode LIKE @search";
                    return View(conn.Query<InventoryItem>(query, new { search = "%" + search + "%" }).ToList());
                }

                return View(conn.Query<InventoryItem>(query).ToList());
            }
        }

        public IActionResult Edit(int id)
        {
            using (var conn = _db.GetConnection())
            {
                conn.Open();
                string query = "SELECT * FROM InventoryItems WHERE Id = @id";
                var item = conn.QueryFirstOrDefault<InventoryItem>(query, new { id });

                if (item == null)
                    return NotFound();

                return View(item);
            }
        }

        [HttpPost]
        public IActionResult Edit(InventoryItem item)
        {
            using (var conn = _db.GetConnection())
            {
                conn.Open();
                string query = @"UPDATE InventoryItems 
                                 SET ItemCode = @ItemCode, ItemName = @ItemName, Quantity = @Quantity, 
                                     Unit = @Unit, Location = @Location 
                                 WHERE Id = @Id";
                conn.Execute(query, item);
            }

            return RedirectToAction("Index");
        }

        public IActionResult Delete(int id)
        {
            using (var conn = _db.GetConnection())
            {
                conn.Open();
                string query = "SELECT * FROM InventoryItems WHERE Id = @id";
                var item = conn.QueryFirstOrDefault<InventoryItem>(query, new { id });

                if (item == null)
                    return NotFound();

                return View(item);
            }
        }

        [HttpPost, ActionName("Delete")]
        public IActionResult DeleteConfirmed(int id)
        {
            using (var conn = _db.GetConnection())
            {
                conn.Open();
                string query = "DELETE FROM InventoryItems WHERE Id = @id";
                conn.Execute(query, new { id });
            }

            return RedirectToAction("Index");
        }

        [HttpGet]
        public IActionResult ExportExcel()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Inventory");

            worksheet.Cell(1, 1).Value = "Kode";
            worksheet.Cell(1, 2).Value = "Nama";
            worksheet.Cell(1, 3).Value = "Qty";
            worksheet.Cell(1, 4).Value = "Satuan";
            worksheet.Cell(1, 5).Value = "Lokasi";

            using var conn = _db.GetConnection();
            conn.Open();
            var items = conn.Query<InventoryItem>("SELECT * FROM InventoryItems");

            int row = 2;
            foreach (var item in items)
            {
                worksheet.Cell(row, 1).Value = item.ItemCode;
                worksheet.Cell(row, 2).Value = item.ItemName;
                worksheet.Cell(row, 3).Value = item.Quantity;
                worksheet.Cell(row, 4).Value = item.Unit;
                worksheet.Cell(row, 5).Value = item.Location;
                row++;
            }

            var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Inventory.xlsx");
        }

        public IActionResult ExportPdf()
        {
            var items = new List<InventoryItem>();

            using (var conn = _db.GetConnection())
            {
                conn.Open();
                string query = "SELECT * FROM InventoryItems";
                items = conn.Query<InventoryItem>(query).ToList();
            }

            var html = this.RenderViewToString("PdfTemplate", items);

            var doc = new HtmlToPdfDocument()
            {
                GlobalSettings = {
                    PaperSize = PaperKind.A4,
                    Orientation = Orientation.Portrait,
                    DocumentTitle = "Inventory PDF"
                },
                Objects = {
                    new ObjectSettings() {
                        HtmlContent = html,
                        WebSettings = { DefaultEncoding = "utf-8" }
                    }
                }
            };

            var pdf = _converter.Convert(doc);
            return File(pdf, "application/pdf", "Inventory.pdf");
        }
    }
}
