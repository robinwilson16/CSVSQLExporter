using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using WinSCP;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;

namespace CSVSQLExporter
{
    class Program
    {
        static async Task<int> Main(string[] args)
        {
            Console.WriteLine("\nExport SQL Table or View to CSV File");
            Console.WriteLine("=========================================\n");

            string? productVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString();
            Console.WriteLine($"Version {productVersion}");
            Console.WriteLine($"Copyright Robin Wilson");

            string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
            string? customConfigFile = null;
            if (args.Length >= 1)
            {
                customConfigFile = args[0];
            }

            if (!string.IsNullOrEmpty(customConfigFile))
            {
                configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, customConfigFile);
            }

            Console.WriteLine($"\nUsing Config File {configFile}");

            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile(configFile, optional: false);

            IConfiguration config;
            try
            {
                config = builder.Build();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e);
                return 1;
            }

            Console.WriteLine($"\nSetting Locale To {config["Locale"]}");

            //Set locale to ensure dates and currency are correct
            CultureInfo culture = new CultureInfo(config["Locale"] ?? "en-GB");
            Thread.CurrentThread.CurrentCulture = culture;
            Thread.CurrentThread.CurrentUICulture = culture;
            CultureInfo.DefaultThreadCurrentCulture = culture;
            CultureInfo.DefaultThreadCurrentUICulture = culture;

            var databaseConnection = config.GetSection("DatabaseConnection");
            var databaseTable = config.GetSection("DatabaseTable");
            var csvFile = config.GetSection("CSVFile");
            var ftpConnection = config.GetSection("FTPConnection");

            var sqlConnection = new SqlConnectionStringBuilder
            {
                DataSource = databaseConnection["Server"],
                UserID = databaseConnection["Username"],
                Password = databaseConnection["Password"],
                InitialCatalog = databaseConnection["Database"],
                TrustServerCertificate = true
            };

            var connectionString = sqlConnection.ConnectionString;

            //Database Connection
            Console.WriteLine("Connecting to Database\n");
            await using var connection = new SqlConnection(connectionString);
            string csvFilePath = "FILE";
            StreamWriter csvFileContents = new StreamWriter(csvFilePath);

            string? columnNameAsFileNameValue = null;
            int? columnNameAsFileNameIndex = null;

            try
            {
                await connection.OpenAsync();
                Console.WriteLine($"\nConnected to {sqlConnection.DataSource}");

                string sql = "";

                if (databaseTable["StoredProcedureCommand"]?.Length > 0)
                {
                    Console.WriteLine($"Executing Stored Procedure {databaseTable["StoredProcedureCommand"]}");

                    sql = $@"[{databaseTable["Database"]}].[{databaseTable["Schema"]}].[{databaseTable["StoredProcedureCommand"]}]";
                }
                else
                {
                    Console.WriteLine($"Loading data from table {databaseTable["TableOrView"]}");

                    sql =
                        $@"SELECT *
                        FROM [{databaseTable["Database"]}].[{databaseTable["Schema"]}].[{databaseTable["TableOrView"]}]";
                }

                await using var command = new SqlCommand(sql, connection);

                //If stored procedure specified with parameters then add these
                if (databaseTable["StoredProcedureCommand"]?.Length > 0)
                {
                    command.CommandType = CommandType.StoredProcedure;

                    if (databaseTable["StoredProcedureParam1IntegerName"]?.Length > 0)
                    {
                        command.Parameters.AddWithValue("@" + databaseTable["StoredProcedureParam1IntegerName"], SqlDbType.Int).Value = databaseTable["StoredProcedureParam1IntegerValue"];
                    }
                    if (databaseTable["StoredProcedureParam2IntegerName"]?.Length > 0)
                    {
                        command.Parameters.AddWithValue("@" + databaseTable["StoredProcedureParam2IntegerName"], SqlDbType.Int).Value = databaseTable["StoredProcedureParam2IntegerValue"];
                    }
                    if (databaseTable["StoredProcedureParam1StringName"]?.Length > 0)
                    {
                        command.Parameters.AddWithValue("@" + databaseTable["StoredProcedureParam1StringName"], SqlDbType.NVarChar).Value = databaseTable["StoredProcedureParam1StringValue"];
                    }
                    if (databaseTable["StoredProcedureParam2StringName"]?.Length > 0)
                    {
                        command.Parameters.AddWithValue("@" + databaseTable["StoredProcedureParam2StringName"], SqlDbType.NVarChar).Value = databaseTable["StoredProcedureParam2StringValue"];
                    }
                }

                await using var reader = await command.ExecuteReaderAsync();

                //Save file to CSV
                Console.WriteLine("\nLoading Data into CSV");

                StringBuilder csvData = new StringBuilder();

                int rowIndex = 0;
                while (await reader.ReadAsync())
                {
                    if (rowIndex == 0 && csvFile.GetValue<bool?>("IncludeHeaders", false) == true)
                    {
                        //Get column names
                        var columnNames = 
                            Enumerable.Range(0, reader.FieldCount)
                                .Select(reader.GetName)
                                .ToList();            

                        //Get file name from column if specified and found
                        for (int cell = 0; cell < reader.FieldCount; cell++)
                        {
                            //Get file name from column if specified and found
                            if (reader.GetName(cell) == csvFile["ColumnNameAsFileName"])
                            {
                                columnNameAsFileNameValue = reader.GetString(cell);
                                columnNameAsFileNameIndex = cell;
                                Console.WriteLine($"Using Custom File Name from Table Column '{csvFile["ColumnNameAsFileName"]}': {columnNameAsFileNameValue}");

                                if (columnNames.Contains(csvFile["ColumnNameAsFileName"]))
                                {
                                    columnNames.Remove(csvFile["ColumnNameAsFileName"]);
                                }
                            }
                        }

                        //Add headers to file data
                        csvData.Append(string.Join(csvFile.GetValue<char>("Delimiter", ','), columnNames));

                        //Append Line
                        csvData.AppendLine();
                    }

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        //Skip column used for file name
                        if (i == columnNameAsFileNameIndex)
                        {
                            continue;
                        }

                        string? value = reader[i].ToString()?.Trim();
                        if (value != null)
                        {
                            //If field contains " then ensure it is double-quoted
                            value = value.Replace("\"", "\"\"");
                            
                            //If field value contains either a comma (or the delimiter from the config file), a double speechmark or a new line character then need to wrap in ""
                            if (csvFile.GetValue<bool?>("AlwaysWrapInSpeechmarks", false) == true)
                            {
                                value = "\"" + value + "\"";
                            }
                            else if (value.Contains(csvFile.GetValue<char>("Delimiter", ',')))
                            {
                                value = "\"" + value + "\"";
                            }
                            else if (value.Contains("\""))
                            {
                                value = "\"" + value + "\"";
                            }
                            else if (value.Contains(Environment.NewLine))
                            {
                                value = "\"" + value + "\"";
                            }

                            csvData.Append(value.Replace(Environment.NewLine, " ") + csvFile.GetValue<char>("Delimiter", ','));
                        }
                    }
                    csvData.Length--; // Remove the last comma
                    csvData.AppendLine();

                    rowIndex++;
                }

                //Close connection
                await connection.CloseAsync();

                Console.WriteLine("\nSaving CSV file");

                string[]? filePaths = { @csvFile["Folder"] ?? "", csvFile["FileName"] ?? "" };

                //If column name specified then use this as the file name instead of the one in the config file
                if (columnNameAsFileNameValue?.Length > 0)
                {
                    if (columnNameAsFileNameValue.Substring(columnNameAsFileNameValue.Length - 4) != ".csv")
                    {
                        columnNameAsFileNameValue = columnNameAsFileNameValue + ".csv";
                    }
                    filePaths = [@csvFile["Folder"] ?? "", columnNameAsFileNameValue ?? ""];
                }

                csvFilePath = Path.Combine(filePaths);
                csvFileContents = new StreamWriter(csvFilePath);

                await csvFileContents.WriteAsync(csvData.ToString());
                csvFileContents.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());

                if (connection != null)
                {
                    await connection.CloseAsync();
                }

                csvFileContents.Close();

                return 1;
            }

            if (System.IO.File.Exists(csvFilePath))
            {

                if (ftpConnection.GetValue<bool?>("UploadFile", false) == true)
                {
                    // Setup session options
                    SessionOptions sessionOptions = new SessionOptions
                    {
                        HostName = ftpConnection["Server"],
                        PortNumber = ftpConnection.GetValue<int>("Port", 21),
                        UserName = ftpConnection["Username"],
                        Password = ftpConnection["Password"]
                    };

                    switch (ftpConnection?["Type"])
                    {
                        case "FTP":
                            sessionOptions.Protocol = Protocol.Ftp;
                            break;
                        case "FTPS":
                            sessionOptions.Protocol = Protocol.Ftp;
                            sessionOptions.FtpSecure = FtpSecure.Explicit;
                            sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = true;
                            break;
                        case "SFTP":
                            sessionOptions.Protocol = Protocol.Sftp;
                            sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = true;
                            break;
                        case "SCP":
                            sessionOptions.Protocol = Protocol.Scp;
                            sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = true;
                            break;
                        default:
                            sessionOptions.Protocol = Protocol.Ftp;
                            break;
                    }

                    if (ftpConnection?["SSHHostKeyFingerprint"]?.Length > 0)
                    {
                        sessionOptions.SshHostKeyFingerprint = ftpConnection["SSHHostKeyFingerprint"];
                        sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = false;
                    }

                    switch (ftpConnection?["Mode"])
                    {
                        case "Active":
                            sessionOptions.FtpMode = FtpMode.Active;
                            break;
                        case "Passive":
                            sessionOptions.FtpMode = FtpMode.Passive;
                            break;
                        default:
                            sessionOptions.FtpMode = FtpMode.Passive;
                            break;
                    }

                    Console.WriteLine("\nUploading CSV File");
                    Console.WriteLine($"Uploading File to {sessionOptions.HostName}");

                    string uploadPath = Path.Combine("/", ftpConnection?["FolderPath"] ?? "");

                    if (uploadPath.Substring(uploadPath.Length - 1) != "/")
                    {
                        uploadPath = uploadPath + "/";
                    }

                    try
                    {
                        using (Session session = new Session())
                        {
                            //When publishing to a self-contained exe file need to specify the location of WinSCP.exe
                            session.ExecutablePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WinSCP.exe");

                            // Connect
                            session.Open(sessionOptions);

                            // Upload files
                            TransferOptions transferOptions = new TransferOptions();
                            transferOptions.TransferMode = TransferMode.Binary;

                            TransferOperationResult transferResult;
                            transferResult =
                                session.PutFiles(csvFilePath, uploadPath, false, transferOptions);

                            // Throw on any error
                            transferResult.Check();

                            // Print results
                            foreach (TransferEventArgs transfer in transferResult.Transfers)
                            {
                                Console.WriteLine("Upload of {0} succeeded", transfer.FileName);
                            }
                        }

                        Console.WriteLine($"File Uploaded to {sessionOptions.HostName} to {uploadPath + columnNameAsFileNameValue ?? csvFile["FileName"] ?? ""}");
                        return 0;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error: {0}", e);
                        return 1;
                    }
                }
                else
                {
                    Console.WriteLine($"Not Uploading File to FTP as Option in Config is False");
                    return 0;
                }
            }
            else
            {
                Console.WriteLine($"The File at {csvFilePath} Could Not Be Found");
                return 1;
            }
        }
    }
}