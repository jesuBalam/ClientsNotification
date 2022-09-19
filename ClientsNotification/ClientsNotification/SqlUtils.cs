using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ClientsNotification
{
    class SqlUtils
    {

        public static DataTable result = new DataTable();
        public static DataTable reference = new DataTable();
        public static DataTable missingStations = new DataTable();

        public static DataTable finalData = new DataTable();

        public static string queryCatalogue = "";
        public static string queryData = "";
        public static string initialDate = "01-08-2022";
        public static string endDate = "07-08-2022";

        public static string initialDateRetro;
        public static string endDateRetro;

        public static int retriesCount = 0;
        private static int countDays = 0;

        public static void Check()
        {
            //Console.WriteLine("PARSING");
            initialDate = ConfigurationManager.AppSettings["initialDate"].ToString();
            endDate = ConfigurationManager.AppSettings["endDate"].ToString();
            initialDateRetro = ConfigurationManager.AppSettings["initialDate"].ToString();
            endDateRetro = ConfigurationManager.AppSettings["endDate"].ToString();

            finalData.Clear();
            result.Clear();
            missingStations.Rows.Clear();
            finalData = new DataTable();
            result = new DataTable();

            queryCatalogue = $"set language spanish " +
                             $"declare @FchIni date ,@FchFin date " +
                             $"set @FchIni = '{initialDate}' set @FchFin = '{endDate}' " +
                             $"SELECT convert(decimal, T1.U_Destino) Destino, T1.[PrcName] Estación, count (T1.U_Destino) Cantidad, T2.ItemName " +
                             $"FROM RHINO_OIL.dbo.ORDR T0 " +
                             $"INNER JOIN RHINO_OIL.dbo.OPRC T1 ON T1.PrcCode = T0.U_Estacion " +
                             $"INNER JOIN RHINO_OIL.dbo.OITM T2 ON T2.ItemCode = T0.U_TipoC " +
                             $"WHERE T0.[DocDate] BETWEEN @FchIni AND @FchFin AND U_Estado <> 'CANCELADO' " +
                             $"group by  T1.U_Destino, T1.[PrcName], T2.ItemName order by convert(decimal, T1.U_Destino)";

            queryData = $"set language spanish " +
                        $"declare @FchIni date ,@FchFin date " +
                        $"set @FchIni = '{initialDate}' set @FchFin = '{endDate}' " +
                        $"SELECT T4.Aux, T0.[U_TAR], T3.Name, convert(date, T0.[DocDate]), T1.U_Destino, T1.[PrcName], " +
                        $"case T2.ItemName " +
                        $"when 'Regular (Mayor a 87 octanos)' then 'PxMagna - 32011' " +
                        $"when 'Premium (Mayor a 91 octanos)' then 'PxPremium - 32012'  " +
                        $"when 'Diésel Automotriz (3)' then 'PxDiesel - 34006' " +
                        $"else '' " +
                        $"end,  " +
                        $"T0.[U_Turno], T0.U_Sel_Vehiculo " +
                        $"FROM[RHINO_OIL].DBO.ORDR T0 " +
                        $"INNER JOIN[RHINO_OIL].DBO.OPRC T1 ON T1.PrcCode = T0.U_Estacion " +
                        $"INNER JOIN[RHINO_OIL].DBO.OITM T2 ON T2.ItemCode = T0.U_TipoC " +
                        $"INNER JOIN[RHINO_OIL].DBO.[@NO_TAR] T3 ON T3.Code = T0.U_TAR " +
                        $"INNER JOIN REPORTES.dbo.AsignacionRhino T4 " +
                        $"ON T0.U_TAR = T4.U_TAR COLLATE Modern_Spanish_CI_AS " +
                        $"AND DATEPART(WEEKDAY, T0.DocDate ) = T4.Dia  " +
                        $"WHERE T0.[DocDate] BETWEEN @FchIni AND @FchFin AND U_Estado<> 'CANCELADO' AND T0.DocStatus = 'O' " +
                        $"order by T4.Aux, T0.[DocDate], T0.U_TAR, convert(decimal, T1.U_Destino)";

            SqlConnection connection = new SqlConnection(string.Format("Data Source={0};database={1}; User ID={2};Password={3}", ConfigurationManager.AppSettings["ServerDatabase"], ConfigurationManager.AppSettings["Database"], ConfigurationManager.AppSettings["User"], ConfigurationManager.AppSettings["Password"]));
            connection.Open();
            using (var command = new SqlCommand(queryCatalogue, connection))
            {
                var adapter = new SqlDataAdapter(command);
                var dataset = new DataSet();
                adapter.Fill(dataset);
                result = dataset.Tables[0];
                Console.WriteLine("Completed query catalogue");
            }

            using (var command = new SqlCommand(queryData, connection))
            {
                var adapter = new SqlDataAdapter(command);
                var dataset = new DataSet();
                adapter.Fill(dataset);
                finalData = dataset.Tables[0];
                Console.WriteLine("Completed query data");
            }
            connection.Close();


            foreach(DataRow row in reference.Rows)
            {
                var rowFounded = (from DataRow dr in result.Rows
                                  where dr["Estación"].ToString().ToUpper().Trim().Replace(" ","") == row["ESTACIONES"].ToString().ToUpper().Trim().Replace(" ", "")
                                  select dr).FirstOrDefault();

                if (rowFounded == null)
                {
                    missingStations.ImportRow(row);
                    Console.WriteLine("MISSING:: " + row["ESTACIONES"]);
                }
            }

            if(missingStations.Rows.Count>0)
            {
                retriesCount++;
                if(retriesCount>=Convert.ToInt32(ConfigurationManager.AppSettings["MaxRetries"].ToString()))
                {
                    FindRetroactiveDates();
                }
                else
                {
                    List<string> emails = new List<string>();
                    foreach (DataRow row in missingStations.Rows)
                    {
                        emails.Add(row["E-MAIL"].ToString().Trim());
                    }
                    EmailUtils.EmailSenderByPackage(emails);
                    WaitUntilNextCheck();
                }
            }
            else
            {
                GenerateFinalExcel();
            }
        }

        public static void WaitUntilNextCheck()
        {
            Console.WriteLine("*** Waiting next check *** ");
            Thread.Sleep(60 * 60 * 1000 * Convert.ToInt32(ConfigurationManager.AppSettings["HourRateReminder"].ToString()));
            Check();
        }

        public static void FindRetroactiveDates()
        {
            countDays++;
            if (DateTime.TryParse(initialDateRetro, out DateTime dateInit))
            {
                initialDateRetro = dateInit.AddDays(-7).ToString("dd-MM-yyyy");
            }

            if (DateTime.TryParse(endDateRetro, out DateTime dateEnd))
            {
                endDateRetro = dateEnd.AddDays(-7).ToString("dd-MM-yyyy");
            }

            string query = $"set language spanish " +
                        $"declare @FchIni date ,@FchFin date " +
                        $"set @FchIni = '{initialDateRetro}' set @FchFin = '{endDateRetro}' " +
                        $"SELECT T4.Aux, T0.[U_TAR], T3.Name, convert(date, T0.[DocDate]), T1.U_Destino, T1.[PrcName], " +
                        $"case T2.ItemName " +
                        $"when 'Regular (Mayor a 87 octanos)' then 'PxMagna - 32011' " +
                        $"when 'Premium (Mayor a 91 octanos)' then 'PxPremium - 32012'  " +
                        $"when 'Diésel Automotriz (3)' then 'PxDiesel - 34006' " +
                        $"else '' " +
                        $"end,  " +
                        $"T0.[U_Turno], T0.U_Sel_Vehiculo " +
                        $"FROM[RHINO_OIL].DBO.ORDR T0 " +
                        $"INNER JOIN[RHINO_OIL].DBO.OPRC T1 ON T1.PrcCode = T0.U_Estacion " +
                        $"INNER JOIN[RHINO_OIL].DBO.OITM T2 ON T2.ItemCode = T0.U_TipoC " +
                        $"INNER JOIN[RHINO_OIL].DBO.[@NO_TAR] T3 ON T3.Code = T0.U_TAR " +
                        $"INNER JOIN REPORTES.dbo.AsignacionRhino T4 " +
                        $"ON T0.U_TAR = T4.U_TAR COLLATE Modern_Spanish_CI_AS " +
                        $"AND DATEPART(WEEKDAY, T0.DocDate ) = T4.Dia  " +
                        $"WHERE T0.[DocDate] BETWEEN @FchIni AND @FchFin AND U_Estado<> 'CANCELADO' AND T0.DocStatus = 'O' " +
                        $"order by T4.Aux, T0.[DocDate], T0.U_TAR, convert(decimal, T1.U_Destino)";

            SqlConnection connection = new SqlConnection(string.Format("Data Source={0};database={1}; User ID={2};Password={3}", ConfigurationManager.AppSettings["ServerDatabase"], ConfigurationManager.AppSettings["Database"], ConfigurationManager.AppSettings["User"], ConfigurationManager.AppSettings["Password"]));
            connection.Open();
            DataTable dataRetroactive = new DataTable();
            using (var command = new SqlCommand(query, connection))
            {
                var adapter = new SqlDataAdapter(command);
                var dataset = new DataSet();
                adapter.Fill(dataset);
                dataRetroactive = dataset.Tables[0];
            }
            connection.Close();

            foreach(DataRow row in dataRetroactive.Rows)
            {
                foreach(DataRow rowM in missingStations.Rows)
                {
                    if (row["U_Destino"].ToString().Trim() == rowM["DESTINO"].ToString().Trim())
                    {
                        if (DateTime.TryParse(initialDateRetro, out DateTime newDate))
                        {
                            int numDays = 7 * countDays;
                            row["Column1"] = newDate.AddDays(numDays).ToString("dd-MM-yyyy");
                        }                        
                        finalData.ImportRow(row);
                    }
                }
            }

            List<string> stationsFounded = new List<string>();
            //Find founded stations to delete
            foreach (DataRow row in missingStations.Rows)
            {
                var rowFounded = (from DataRow dr in dataRetroactive.Rows
                                  where dr["U_Destino"].ToString().ToUpper().Trim().Replace(" ", "") == row["DESTINO"].ToString().ToUpper().Trim().Replace(" ", "")
                                  select dr).FirstOrDefault();

                if (rowFounded != null)
                {

                    stationsFounded.Add(row["DESTINO"].ToString());
                    //new Sender email
                    if (!stationsFounded.Contains(row["DESTINO"].ToString()))
                    {
                        List<string> emails = new List<string>();
                        emails.Add(row["E-MAIL"].ToString().Trim());
                        EmailUtils.EmailSenderByPackage(emails, "", $"Se envió información retroactiva de tu reporte con la fecha: {rowFounded["Column1"]}");
                    }
                    Console.WriteLine(row["ESTACIONES"] + "founded in " + initialDateRetro + " : " + endDateRetro);
                }
            }

            for (int i = missingStations.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr = missingStations.Rows[i];
                string found = "";
                found = stationsFounded.Find(element => element == missingStations.Rows[i]["DESTINO"].ToString());
                if (!string.IsNullOrEmpty(found))
                {
                    dr.Delete();
                }
            }

            if (missingStations.Rows.Count > 0)
            {
                Console.WriteLine($"Try to find last week. ({initialDateRetro} : { endDateRetro})" + missingStations.Rows.Count + " stations left");
                Thread.Sleep(3 * 1000 );
                FindRetroactiveDates();
            }
            else
            {
                GenerateFinalExcel();
            }
        }


        public static void GenerateTableFromExcel(string path = @"C:\Users\enriq\Downloads\Destinos.xlsx")
        {
            using (ExcelPackage package = new ExcelPackage(ConfigurationManager.AppSettings["PathReference"].ToString()))//"C:/Users/ControlGas/Desktop/Destinos.xlsx"
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    reference.Columns.Add(firstRowCell.Text);
                    missingStations.Columns.Add(firstRowCell.Text);
                }                
                for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                    var row = reference.NewRow();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }

                    reference.Rows.Add(row);
                }
            }
            Console.WriteLine("Reference readed");
        }

        public static void GenerateFinalExcel(string path = "C:/Users/ControlGas/Desktop/FinalData.xlsx")
        {
            Console.WriteLine("Generating final excel");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //Final path must not exist
            using (ExcelPackage package = new ExcelPackage(ConfigurationManager.AppSettings["PathFinalExcel"].ToString()))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells["A1"].LoadFromDataTable(finalData, true);
                package.Save();
            }
            List<string> emails = new List<string>();
            foreach (DataRow row in reference.Rows)
            {
                emails.Add(row["E-MAIL"].ToString().Trim());
            }
            EmailUtils.EmailSenderByPackage(emails, ConfigurationManager.AppSettings["PathFinalExcel"].ToString());
        }
    }
}
