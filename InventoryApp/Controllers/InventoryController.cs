using Microsoft.AspNetCore.Mvc;
using InventoryApp.Data;
using InventoryApp.Models;
using Microsoft.Data.SqlClient;
using System.IO;
using ClosedXML.Excel;
using DinkToPdf.Contracts;
using DinkToPdf;
using InventoryApp.Helpers;

namespace InventoryApp.Controllers
{
    public class InventoryController : Controller
    {
        // Define connection
        private readonly Db _db;
        private readonly IConverter _converter;

        public InventoryController(Db db, IConverter converter)
        {
            _db = db;
            _converter = converter;
        }

        // GET: Inventory/Create
        public IActionResult Create()
        {
            if (HttpContext.Session.GetString("role") != "admin")
                return Unauthorized();

            return View();
        }

        // POST: Inventory/Create
        [HttpPost]
        public IActionResult Create(InventoryItem item)
        {
            using (var conn = _db.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"INSERT INTO InventoryItems 
            (ItemCode, ItemName, Quantity, Unit, Location) 
            VALUES (@code, @name, @qty, @unit, @loc)", conn);

                cmd.Parameters.AddWithValue("@code", item.ItemCode);
                cmd.Parameters.AddWithValue("@name", item.ItemName);
                cmd.Parameters.AddWithValue("@qty", item.Quantity);
                cmd.Parameters.AddWithValue("@unit", item.Unit);
                cmd.Parameters.AddWithValue("@loc", item.Location);

                cmd.ExecuteNonQuery();
            }

            return RedirectToAction("Index");
        }

        // GET: List
        public IActionResult Index(string search)
        {
            var items = new List<InventoryItem>();

            using (var conn = _db.GetConnection())
            {
                conn.Open();

                string query = "SELECT * FROM InventoryItems";
                if (!string.IsNullOrEmpty(search))
                {
                    query += " WHERE ItemName LIKE @search OR ItemCode LIKE @search";
                }

                var cmd = new SqlCommand(query, conn);

                if (!string.IsNullOrEmpty(search))
                {
                    cmd.Parameters.AddWithValue("@search", "%" + search + "%");
                }

                var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    items.Add(new InventoryItem
                    {
                        Id = (int)reader["Id"],
                        ItemCode = reader["ItemCode"].ToString(),
                        ItemName = reader["ItemName"].ToString(),
                        Quantity = (int)reader["Quantity"],
                        Unit = reader["Unit"].ToString(),
                        Location = reader["Location"].ToString(),
                        CreatedAt = (DateTime)reader["CreatedAt"]
                    });
                }
            }

            return View(items);
        }


        // GET: Inventory/Edit/
        public IActionResult Edit(int id)
        {
            InventoryItem item = null;

            using (var conn = _db.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("SELECT * FROM InventoryItems WHERE Id = @id", conn);
                cmd.Parameters.AddWithValue("@id", id);

                var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    item = new InventoryItem
                    {
                        Id = (int)reader["Id"],
                        ItemCode = reader["ItemCode"].ToString(),
                        ItemName = reader["ItemName"].ToString(),
                        Quantity = (int)reader["Quantity"],
                        Unit = reader["Unit"].ToString(),
                        Location = reader["Location"].ToString(),
                        CreatedAt = (DateTime)reader["CreatedAt"]
                    };
                }
            }

            if (item == null)
                return NotFound();

            return View(item);
        }

        // POST: Inventory/Edit/
        [HttpPost]
        public IActionResult Edit(InventoryItem item)
        {
            using (var conn = _db.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"UPDATE InventoryItems 
            SET ItemCode = @code, ItemName = @name, Quantity = @qty, 
                Unit = @unit, Location = @loc 
            WHERE Id = @id", conn);

                cmd.Parameters.AddWithValue("@id", item.Id);
                cmd.Parameters.AddWithValue("@code", item.ItemCode);
                cmd.Parameters.AddWithValue("@name", item.ItemName);
                cmd.Parameters.AddWithValue("@qty", item.Quantity);
                cmd.Parameters.AddWithValue("@unit", item.Unit);
                cmd.Parameters.AddWithValue("@loc", item.Location);

                cmd.ExecuteNonQuery();
            }

            return RedirectToAction("Index");
        }

        // GET: Inventory/Delete/
        public IActionResult Delete(int id)
        {
            InventoryItem item = null;

            using (var conn = _db.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("SELECT * FROM InventoryItems WHERE Id = @id", conn);
                cmd.Parameters.AddWithValue("@id", id);
                var reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    item = new InventoryItem
                    {
                        Id = (int)reader["Id"],
                        ItemCode = reader["ItemCode"].ToString(),
                        ItemName = reader["ItemName"].ToString(),
                        Quantity = (int)reader["Quantity"],
                        Unit = reader["Unit"].ToString(),
                        Location = reader["Location"].ToString(),
                        CreatedAt = (DateTime)reader["CreatedAt"]
                    };
                }
            }

            if (item == null)
                return NotFound();

            return View(item);
        }

        // POST: Inventory/Delete/
        [HttpPost, ActionName("Delete")]
        public IActionResult DeleteConfirmed(int id)
        {
            using (var conn = _db.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("DELETE FROM InventoryItems WHERE Id = @id", conn);
                cmd.Parameters.AddWithValue("@id", id);
                cmd.ExecuteNonQuery();
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

            var conn = _db.GetConnection();
            conn.Open();
            var cmd = new SqlCommand("SELECT * FROM InventoryItems", conn);
            var reader = cmd.ExecuteReader();

            int row = 2;
            while (reader.Read())
            {
                worksheet.Cell(row, 1).Value = reader["ItemCode"].ToString();
                worksheet.Cell(row, 2).Value = reader["ItemName"].ToString();
                worksheet.Cell(row, 3).Value = Convert.ToInt32(reader["Quantity"]);
                worksheet.Cell(row, 4).Value = reader["Unit"].ToString();
                worksheet.Cell(row, 5).Value = reader["Location"].ToString();
                row++;
            }

            var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Inventory.xlsx");
        }

        

        //public IActionResult ExportPdf()
        //{
        //    var items = new List<InventoryItem>();

        //    using (var conn = _db.GetConnection())
        //    {
        //        conn.Open();
        //        var cmd = new SqlCommand("SELECT * FROM InventoryItems", conn);
        //        var reader = cmd.ExecuteReader();

        //        while (reader.Read())
        //        {
        //            items.Add(new InventoryItem
        //            {
        //                Id = (int)reader["Id"],
        //                ItemCode = reader["ItemCode"].ToString(),
        //                ItemName = reader["ItemName"].ToString(),
        //                Quantity = (int)reader["Quantity"],
        //                Unit = reader["Unit"].ToString(),
        //                Location = reader["Location"].ToString(),
        //                CreatedAt = (DateTime)reader["CreatedAt"]
        //            });
        //        }
        //    }

        //    // Render partial view ke string HTML
        //    var html = this.RenderViewToString("PdfTemplate", items);

        //    var doc = new HtmlToPdfDocument()
        //    {
        //        GlobalSettings = {
        //    PaperSize = PaperKind.A4,
        //    Orientation = Orientation.Portrait,
        //    DocumentTitle = "Inventory PDF"
        //},
        //        Objects = {
        //    new ObjectSettings() {
        //        HtmlContent = html,
        //        WebSettings = { DefaultEncoding = "utf-8" }
        //    }
        //}
        //    };

        //    var pdf = _converter.Convert(doc);
        //    return File(pdf, "application/pdf", "Inventory.pdf");
        //}

    }
}
