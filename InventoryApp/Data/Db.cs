using System.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace InventoryApp.Data
{
    public class Db
    {
        private readonly IConfiguration _config;
        private readonly string _connStr;

        public Db(IConfiguration config)
        {
            _config = config;
            _connStr = _config.GetConnectionString("DefaultConnection");
        }

        public SqlConnection GetConnection()
        {
            return new SqlConnection(_connStr);
        }
    }
}