using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace Excel_DB_integration
{
    public static class ConnectionString
    {
        public static SqlConnection connection()
        {
            SqlConnection conn = new SqlConnection("Data Source=CEYHUNKUTAHYALI\\SQLEXPRESS;Initial Catalog=Project_db;Integrated Security=True");
            conn.Open();
            return conn;
        }
    }
}



