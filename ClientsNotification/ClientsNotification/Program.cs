using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClientsNotification
{
    class Program
    {
        static void Main(string[] args)
        {           
            SqlUtils.GenerateTableFromExcel();
            SqlUtils.Check();
            Console.WriteLine("Press any key to exit");
            Console.ReadKey();

            //Check query ---CHECK
            //Construct 2 tables (stations and final data) CHECK
            //Verifiy that stations are same that catalogue reference CHECK
            //Construct table with missing stations CHECK
            // Send mail to missing stations emails (Program every x hours) and verify missing stations again

            // After X retries, Verify missing stations
            //Construct final table if not missing stations CHECK
            //Consult query with retroactive dates to find missing stations
            //Construct final table CHECK
            //Construct excel CHECK
            //Send mail to all emails of catalogue CHECK


            //Enviar correo a email encargado mencionandole que info y fecha se encontro informacion
            //Actualizar fecha de tabla final excel cuando sea retroactiva
        }
    }
}
