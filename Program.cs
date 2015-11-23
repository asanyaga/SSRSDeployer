using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Protocols;
using System.Configuration;
using VirtualCity.SSRSDeployer.Console.ReportService2010;
using Microsoft.SqlServer.Management;
using Microsoft.SqlServer.Server;
using Microsoft.SqlServer.Management.Smo;
using System.Data.SqlClient;
using Microsoft.SqlServer.Management.Common;
namespace VirtualCity.SSRSDeployer.Console
{
    class Program
    {
      
        static void Main(string[] args)
        {
            try
            {
                ExecuteDatabaseScripts();
                System.Console.ReadLine();
            }
            catch(Exception e)
            {
                System.Console.Write(e.StackTrace);
            }
            System.Console.ReadLine();
           
        }

        static void ExecuteDatabaseScripts()
        {
            string sqlConnectionString = "Integrated Security=SSPI;" +
            "Persist Security Info=True;Initial Catalog="+ Properties.Settings.Default.SQLDatabaseName +";Data Source="+Properties.Settings.Default.SQLServerInstance;
            DirectoryInfo di = new DirectoryInfo(Properties.Settings.Default.SQLScriptsFolder);
            FileInfo[] scriptFiles = di.GetFiles("*.sql");
            System.Console.WriteLine("Connecting to Database...." + Properties.Settings.Default.SQLDatabaseName);
            SqlConnection connection = new SqlConnection(sqlConnectionString);
            Server server = new Server(new ServerConnection(connection));
            System.Console.WriteLine("Connected to Database...." + Properties.Settings.Default.SQLDatabaseName);
            List<String> errors = new List<string>();
            foreach (FileInfo fi in scriptFiles)
            {
                try
                {
                    FileInfo fileInfo = new FileInfo(fi.FullName);
                    string script = fileInfo.OpenText().ReadToEnd();

                    System.Console.WriteLine("Executing Script " + fi.FullName);

                    server.ConnectionContext.ExecuteNonQuery(script);
  
                }
                catch(Exception ex)
                {
                    System.Console.WriteLine(ex.StackTrace);
                    errors.Add(ex.Message);
                }
             }
            if (connection.State == System.Data.ConnectionState.Open)
            {
                connection.Close();
            }
            foreach(String error in errors)
            {
                System.Console.WriteLine(error);
            }
    
        }

        static DataSourceDefinition ConfigureDataSourceDefinition()
        {
            DataSourceDefinition dataSourceDefinition = new DataSourceDefinition();

            dataSourceDefinition.CredentialRetrieval = CredentialRetrievalEnum.Integrated;
            dataSourceDefinition.ConnectString = "Data Source="+ Properties.Settings.Default.SQLServerInstance + "Initial Catalog="+Properties.Settings.Default.SQLDatabaseName + "Trusted_Connection=true";
            dataSourceDefinition.Enabled = true;
            dataSourceDefinition.EnabledSpecified = true;
            dataSourceDefinition.Extension = "SQL";
            dataSourceDefinition.ImpersonateUserSpecified = false;
            
            return dataSourceDefinition;

        }
        static String[] GetDataSourceFileNames()
        {
            String[] dataSources = Directory.GetFiles(Properties.Settings.Default.ReportsSourceFolder, "*.rds", SearchOption.TopDirectoryOnly);
            return dataSources;
        }
        static void UploadDataSource(String dataSourceName,DataSourceDefinition definition)
        {
            ReportingService2010 rs = new ReportingService2010();
            rs.Url = Properties.Settings.Default.SSRSUrl;
            rs.Credentials = System.Net.CredentialCache.DefaultCredentials;
            rs.CreateDataSource(dataSourceName, Properties.Settings.Default.SSRSFolder, false, definition, null);

        }
        static void UploadReport()
        {
            ReportService2010.ReportingService2010 rs = new ReportService2010.ReportingService2010();
            string serviceUrl = Properties.Settings.Default.SSRSUrl;
            rs.Credentials = System.Net.CredentialCache.DefaultCredentials;
            rs.Url = serviceUrl;
            string strItemType = "Report";
            string strName = "Sales By Product";
            Byte[] definition = null;
            ReportService2010.Warning[] warnings = null;

            String[] reportFiles = Directory.GetFiles(Properties.Settings.Default.ReportsSourceFolder, "*.rdl", SearchOption.TopDirectoryOnly);

            foreach (String reportFile in reportFiles)
            {
                try
                {
                    System.Console.WriteLine("Reading the report file");

                    FileStream stream = File.OpenRead(reportFile);
                    definition = new Byte[stream.Length];
                    stream.Read(definition, 0, (int)stream.Length);
                    stream.Close();

                    System.Console.WriteLine("Finished Reading the report file");
                }
                catch (IOException e)
                {
                    System.Console.WriteLine(e.Message);
                }

                try
                {
                    string parent = Properties.Settings.Default.SSRSFolder;
                    System.Console.WriteLine("Uploading the report file");
                    ReportService2010.CatalogItem report = rs.CreateCatalogItem(strItemType, strName, parent, true, definition, null, out warnings);

                    if (warnings != null)
                    {
                        foreach (ReportService2010.Warning warning in warnings)
                        {
                            System.Console.WriteLine(warning.Message);
                        }
                    }

                    else
                        System.Console.WriteLine("Report: {0} created successfully " +
                                          " with no warnings", strName);
                    System.Console.ReadLine();
                }
                catch (SoapException e)
                {
                    System.Console.WriteLine(e.Detail.InnerXml.ToString());
                }

            }
        }
    }
}
