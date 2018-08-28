using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2
{
    class Fileload
    {
        string constr = "";

        public List<OleDbDataReader> readers = new List<OleDbDataReader>();
        public Fileload()
        { }
        public List<OleDbDataReader> Getlist()
        {
            return readers;
        }

        public Fileload(string filepath)
        {
            this.constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath + ";Extended Properties='EXCEL 12.0;HDR=YES;';";
        }

        public void Reader()
        {
            OleDbConnection sncon = new OleDbConnection(constr);
            sncon.Open();
            DataTable sheetnamedt = sncon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetname = (string)sheetnamedt.Rows[0][2];
            OleDbCommand cmd = new OleDbCommand
            {
                Connection = sncon,

                // excel data reading
                CommandText = new StringBuilder().AppendFormat("select [Country_code],[AgreementID],[Amp_ID],[ASM],[ASM_EMAIL],CDate([Contract Create Date]) as [Contract Create Date],CDate([Contract End Date]) as [Contract End Date],CDate([Contract Start Date]) as [Contract Start Date],[CONTRACT_ADMIN],[COUNTRY_Name],[Customer],[DASM],[Document_Type],[FUNCTIONAL_LOCATION],[GBU],[Group Name],[Hierarchy_Level_1],[Hierarchy_Level_2],[Hierarchy_Level_3],[Hierarchy_Level_4],IIF(ISNULL(HPPID_Activation_Date),null,CDate(HPPID_Activation_Date)) as HPPID_Activation_Date,[HPPID_EMail],[HW_ShipTo_City],[HW_ShipTo_Contact],[HW_ShipTo_Email],[HW_ShipTo_Phone],[ShipTo_Postal_Code],[Linked_Status],[OBID],[Part_Desc],[Part_Number],[PkgID],[PL],[Purchase_Order_Identifier],[Quantity],[Region_Name],[RST_Eligible],[Sales Order Nbr],[SALES_REP],[Sales_Route],[Serial_Number],[Service_Type],[SGID],[Ship_AMID2],[Ship_AMID2_Ind_Seg_Name],[Ship_AMID2_Name],[Ship_AMID4],[Ship_AMID4_Name],[Sold_AMID2],[Sold_AMID2_Name],[Sold_AMID4],[Sold_AMID4_Name],[Sub_region],[Sub_Region1],[SYS_Manager_Email],[SYS_Manager_Phone],[SYSTEM_Manager_Name],[TAM] from [ {0}A1:BK]", sheetname).ToString()
            };
            OleDbDataReader reader = cmd.ExecuteReader();

            readers.Add(reader);
        }
    }
}
