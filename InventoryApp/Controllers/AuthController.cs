using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using InventoryApp.Data;
using Dapper;

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

        var user = conn.QueryFirstOrDefault<dynamic>(
            "SELECT * FROM Users WHERE Username = @Username AND Password = @Password",
            new { Username = username, Password = password }
        );

        if (user != null)
        {
            // Simpan info user ke session
            HttpContext.Session.SetString("username", (string)user.Username);
            HttpContext.Session.SetString("role", (string)user.Role);

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
