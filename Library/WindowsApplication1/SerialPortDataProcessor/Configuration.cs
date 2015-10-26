using System.Configuration;
using System.Reflection;

namespace SerialPortDataProcessor
{
    public class Configuration
    {

        public static string GetConnectionString()
        {
            string connectionString = ConfigurationManager.AppSettings["ConnectionString"];
            if (string.IsNullOrEmpty(connectionString))
            {
                string filename = Assembly.GetExecutingAssembly().Location;
                System.Configuration.Configuration configuration =
                    ConfigurationManager.OpenExeConfiguration(filename);
                if (configuration != null)
                    connectionString = configuration.AppSettings.Settings["ConnectionString"].Value;
            }
            
            return connectionString;
        }

        public static string GetPrinterName()
        {
            string printerName = ConfigurationManager.AppSettings["PrinterName"];
            if (string.IsNullOrEmpty(printerName))
            {
                string filename = Assembly.GetExecutingAssembly().Location;
                System.Configuration.Configuration configuration =
                    ConfigurationManager.OpenExeConfiguration(filename);
                if (configuration != null)
                    printerName = configuration.AppSettings.Settings["PrinterName"].Value;
            }

            return printerName;
        }
    }
}
