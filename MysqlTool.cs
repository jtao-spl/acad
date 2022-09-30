using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace acad01
{
    class MysqlTool
    {
        private MySqlConnection conn = null;
        public enum ELEMENT_FIRST_TYPE
        {
            SIZED_ELEMENT,
            GEOMETRICAL_TOLERNACE,
            SURFACE_ROUGHNESS,
            OTHER
        }
        public enum ELEMENT_SIZED_ELEMENT_SUB_TYPE
        {
            LINE,
            DIAMETER,
            RADIAL,
            ANGLE,
        }
        public enum ELEMENT_GEO_TOLERANCE_SUB_TYPE
        {

        }
        public enum ELEMENT_SURFACE_ROUGHNESS_SUB_TYPE
        {
            RA
        }
        public MySqlConnection getConnection()
        {
            if (conn != null) {
                return conn;
            }else {
                MySqlConnectionStringBuilder builder = new MySqlConnectionStringBuilder();
                builder.Server = "localhost";
                builder.Port = 3306;
                builder.UserID = "root";
                builder.Password = "KTIpdx91@1";
                conn = new MySqlConnection(builder.ConnectionString);
                return conn;
            
            }
        }
    }

    
  
}
