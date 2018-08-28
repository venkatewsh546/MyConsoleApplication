using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;


namespace ConsoleApplication2
{
    class Dataload
    {
        public string vconstring;
        public string query;
        public string pl;
        public DataTable dt;
        
        public Dataload(string vconstring, string query, string pl, DataTable dt)
        {
           this.vconstring = vconstring;
           this.query = "select * from dbo.CORE_IB_DATA_ISS where Product_Line_Id='" + pl + "'";
           this.pl = pl;
           this.dt = dt;           

        }

        public void Loaddata()
        {
            try
            {
                Console.WriteLine("{0} Thread started ", pl);

                SqlConnection scon = new SqlConnection
                {
                    ConnectionString = vconstring
                };
                scon.Open();

                SqlCommand vcmd = new SqlCommand
                {
                    Connection = scon,
                    CommandText = new StringBuilder().AppendFormat("select * from dbo.CORE_IB_DATA_ISS where Product_Line_Id='{0}'", pl).ToString()
                };
                SqlDataReader reader = vcmd.ExecuteReader();

                string data = "";

                foreach (DataColumn dtc in dt.Columns)
                {
                    data += "\"" + dtc.ColumnName + "\"" + ",";
                }
                File.AppendAllText(@"C:\Users\pothugun\Downloads\globalshare\coreibdata_" + pl + ".csv", data.TrimEnd(',') + System.Environment.NewLine);
                data = "";


                while (reader.Read())
                {

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        data += "\"" + reader[i] + "\"" + ",";
                    }

                    File.AppendAllText(@"C:\Users\pothugun\Downloads\globalshare\coreibdata_" + pl + ".csv", data.TrimEnd(',') + System.Environment.NewLine);
                    data = "";

                }

                Console.WriteLine("{0} Thread Ended ", pl);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
