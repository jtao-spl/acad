using System;
using System.Collections.Generic;
using System.Data;
using MySql.Data.MySqlClient;

namespace acad
{
    class ComponentTool
    {
        private string getConnectionString()
        {
       
            MySqlConnectionStringBuilder builder = new MySqlConnectionStringBuilder();
            uint port;
            builder.Server = APIConfHelper.DBSettings["Server"].ToString();
            try
            {
                port = uint.Parse(APIConfHelper.DBSettings["Port"].ToString());
            }
            catch (FormatException)
            {

                throw new Exception("数据库链接配置异常：Port非整数。");
            }
            builder.Port = port;
            builder.UserID = APIConfHelper.DBSettings["UserID"].ToString();
            builder.Password = APIConfHelper.DBSettings["Password"].ToString();
            builder.Pooling = false;
            builder.Database = APIConfHelper.DBSettings["Database"].ToString();
            return builder.ConnectionString;
        }

        public int executeQueryId(string sql)
        {
            MySqlConnection conn = new MySqlConnection(getConnectionString());
            if (!conn.Ping())
            {
                conn.Open();
            }

            MySqlCommand cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            Object obj = cmd.ExecuteScalar();
            if (conn.State != ConnectionState.Closed)
            {
                conn.Close();
            }
            if(obj == null)
            {
                return -1;
            }
            return (int)obj;
            
        }
        public void executeInsert(string[] sqls)
        {
            MySqlConnection conn = new MySqlConnection(getConnectionString());
            if (!conn.Ping())
            {
                conn.Open();
            }
            MySqlCommand cmd = conn.CreateCommand();
            MySqlTransaction trans = conn.BeginTransaction();
            cmd.Transaction = trans;
            try
            {
                for (int i = 0; i < sqls.Length; i++)
                {
                    cmd.CommandText = sqls[i];
                    int ret_val = cmd.ExecuteNonQuery();
                }
                trans.Commit();
            }
            catch (Exception ex)
            {
                trans.Rollback();
                throw ex;
            }
            finally
            {
                if (conn.State != System.Data.ConnectionState.Closed)
                {
                    conn.Close();
                }
            }
        }
        public int CreateComponent()
        {
            string name = $"零件{ String.Format("{0:yyyyMMddHHmmss}", DateTime.Now)}";
            string sql = $"insert into t_component(ComponentName) values(\"{name}\")";
            List<string> sqls = new List<string>();
            sqls.Add(sql);
            executeInsert(sqls.ToArray());
            int Id = executeQueryId($"select * from t_component where ComponentName=\"{name}\"");
            return Id;
        }

        public void CreateComponentSize(int ComponentId, Element e)
        {
            List<string> sqls = new List<string>();
            for (int i = 0; i < e.sizedElements.Length; i++)
            {
                sqls.Add($"insert into t_component_size" +
                    $"(ComponentId,FirstType,SecondType,BaseSize,UpSize,BottomSize) values" +
                    $"({ComponentId},{(int)ELEMENT_FIRST_TYPE.SIZED_ELEMENT},{ (int)e.sizedElements[i].sizeType}," +
                    $"{e.sizedElements[i].baseSize},{e.sizedElements[i].upperSize},{e.sizedElements[i].lowerSize})");
            }
            for (int i = 0; i < e.geometricalTolerances.Length; i++)
            {
                sqls.Add($"insert into t_component_size" +
                    $"(ComponentId,FirstType,GeoToleranceType,GeoToleranceVal) values(" +
                    $"{ComponentId},{(int)ELEMENT_FIRST_TYPE.GEOMETRICAL_TOLERNACE}," +
                    $"\"{e.geometricalTolerances[i].ToneranceType}\",\"{e.geometricalTolerances[i].TonerancePrecision}\")");
            }
            for (int i = 0; i < e.surfaceRoughnesses.Length; i++)
            {
                sqls.Add($"insert into t_component_size" +
                    $"(ComponentId,FirstType,SurfaceRoughnessType,SurfaceRoughnessVal) values(" +
                    $"{ComponentId},{(int)ELEMENT_FIRST_TYPE.SURFACE_ROUGHNESS}," +
                    $"\"{e.surfaceRoughnesses[i].RoughnessType}\",\"{e.surfaceRoughnesses[i].RoughnessValue}\")");
            }
            for (int i = 0; i < e.otherRequirements.Length; i++)
            {
                sqls.Add($"insert into t_component_size" +
                    $"(ComponentId,FirstType,OtherRequirements) values(" +
                    $"{ComponentId},{(int)ELEMENT_FIRST_TYPE.OTHER},\"{e.otherRequirements[i].requirement}\"" +
                    $")");
            }
            executeInsert(sqls.ToArray());
        }

        public void SaveOriginFile(int ComponentId, string fileName, byte[] content)
        {
            string sql = $"insert into t_component_file(ComponentId, FileName, FileContent)values(" +
                $"@ComponentId, @fileName, @binary)";
            MySqlConnection conn = new MySqlConnection(getConnectionString());
            MySqlCommand cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.Add("@binary", MySqlDbType.Binary, content.Length);
            cmd.Parameters["@binary"].Value = content;
            cmd.Parameters.Add("@ComponentId", MySqlDbType.Int32);
            cmd.Parameters["@ComponentId"].Value = ComponentId;
            cmd.Parameters.Add("@fileName", MySqlDbType.VarChar, 200);
            cmd.Parameters["@fileName"].Value = fileName;

            if (!conn.Ping())
            {
                conn.Open();
            }
            MySqlTransaction trans = conn.BeginTransaction();
            cmd.Transaction = trans;
            try
            {
                int ret_val = cmd.ExecuteNonQuery();
                trans.Commit();
            }
            catch (Exception ex)
            {
                trans.Rollback();
                throw ex;
            }
            finally
            {
                if (conn.State != System.Data.ConnectionState.Closed)
                {
                    conn.Close();
                }
            }
        }
    }


}
