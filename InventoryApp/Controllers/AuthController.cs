using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using InventoryApp.Data;

public class AuthController : Controller
{
    private readonly Db _db;

    public AuthController(Db db)
    {
        _db = db;
    }

    public IActionResult Login()
    {
        return View();
    }

    [HttpPost]
    public IActionResult Login(string username, string password)
    {
        using var conn = _db.GetConnection();
        conn.Open();
        var cmd = new SqlCommand("SELECT * FROM Users WHERE Username = @u AND Password = @p", conn);
        cmd.Parameters.AddWithValue("@u", username);
        cmd.Parameters.AddWithValue("@p", password);

        var reader = cmd.ExecuteReader();
        if (reader.Read())
        {
            HttpContext.Session.SetString("username", reader["Username"].ToString());
            HttpContext.Session.SetString("role", reader["Role"].ToString());

            return RedirectToAction("Index", "Inventory");
        }

        ViewBag.Error = "Username / Password salah!";
        return View();
    }

    public IActionResult Logout()
    {
        HttpContext.Session.Clear();
        return RedirectToAction("Login");
    }
}
