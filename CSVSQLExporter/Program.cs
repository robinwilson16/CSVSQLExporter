﻿using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System.Text;
using WinSCP;

namespace CSVSQLExporter
{
    class Program
    {
        static async Task<int> Main(string[] args)
        {
            Console.WriteLine("\nExport SQL Table or View to CSV File");
            Console.WriteLine("=========================================\n");
            Console.WriteLine("Copyright Robin Wilson");

            string configFile = "appsettings.json";
            string? customConfigFile = null;
            if (args.Length >= 1)
            {
                customConfigFile = args[0];
            }

            if (!string.IsNullOrEmpty(customConfigFile))
            {
                configFile = customConfigFile;
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

            var databaseConnection = config.GetSection("DatabaseConnection");
            var databaseTable = config.GetSection("DatabaseTable");
            var csvFile = config.GetSection("CSVFile");
            var ftpConnection = config.GetSection("FTPConnection");
            string[]? filePaths = { @csvFile["Folder"] ?? "", csvFile["FileName"] ?? "" };
            string csvFilePath = Path.Combine(filePaths);

            var sqlConnection = new SqlConnectionStringBuilder
            {
                DataSource = databaseConnection["Server"],
                UserID = databaseConnection["Username"],
                Password = databaseConnection["Password"],
                InitialCatalog = databaseConnection["Database"],
                TrustServerCertificate = true
            };

            var connectionString = sqlConnection.ConnectionString;

            try
            {
                //Database Connection
                await using var connection = new SqlConnection(connectionString);
                Console.WriteLine("Connecting to Database\n");

                await connection.OpenAsync();
                Console.WriteLine($"\nConnected to {sqlConnection.DataSource}");
                Console.WriteLine($"Loading data from table {databaseTable["TableOrView"]}");

                var sql =
                    @"SELECT *
                    FROM [" + databaseTable["Database"] + "].[" + databaseTable["Schema"] + "].[" + databaseTable["TableOrView"] + "]";

                await using var command = new SqlCommand(sql, connection);
                await using var reader = await command.ExecuteReaderAsync();


                //Save file to CSV
                Console.WriteLine("Loading Data into CSV\n");

                StringBuilder csvData = new StringBuilder();
                StreamWriter csvFileContents = new StreamWriter(csvFilePath);

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

                        //Add headers to file data
                        csvData.Append(string.Join(",", columnNames));

                        //Append Line
                        csvData.AppendLine();
                    }
                    
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        string? value = reader[i].ToString();
                        if (value != null)
                        {
                            //If field contains " then ensure it is double-quoted
                            value = value.Replace("\"", "\"\"");
                            
                            //If field value contains either a comma or a new line character then need to wrap in ""
                            if (value.Contains(","))
                            {
                                value = "\"" + value + "\"";
                            }
                            else if (value.Contains(Environment.NewLine))
                            {
                                value = "\"" + value + "\"";
                            }

                            csvData.Append(value.Replace(Environment.NewLine, " ") + ",");
                        }
                    }
                    csvData.Length--; // Remove the last comma
                    csvData.AppendLine();

                    rowIndex++;
                }

                //Close connection
                connection.Close();
                csvFileContents.Write(csvData.ToString());
                csvFileContents.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
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

                    switch (ftpConnection["Type"])
                    {
                        case "FTP":
                            sessionOptions.Protocol = Protocol.Ftp;
                            break;
                        case "FTPS":
                            sessionOptions.Protocol = Protocol.Ftp;
                            sessionOptions.FtpSecure = FtpSecure.Explicit;
                            break;
                        case "SFTP":
                            sessionOptions.Protocol = Protocol.Sftp;
                            break;
                        default:
                            sessionOptions.Protocol = Protocol.Ftp;
                            break;
                    }

                    switch (ftpConnection["Mode"])
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

                    Console.WriteLine($"Uploading File to {sessionOptions.HostName}");

                    try
                    {
                        using (Session session = new Session())
                        {
                            //When publishing to a self-contained exe file need to specify the location of WinSCP.exe
                            session.ExecutablePath = AppDomain.CurrentDomain.BaseDirectory + "\\WinSCP.exe";

                            // Connect
                            session.Open(sessionOptions);

                            // Upload files
                            TransferOptions transferOptions = new TransferOptions();
                            transferOptions.TransferMode = TransferMode.Binary;

                            TransferOperationResult transferResult;
                            transferResult =
                                session.PutFiles(csvFilePath, "/", false, transferOptions);

                            // Throw on any error
                            transferResult.Check();

                            // Print results
                            foreach (TransferEventArgs transfer in transferResult.Transfers)
                            {
                                Console.WriteLine("Upload of {0} succeeded", transfer.FileName);
                            }
                        }

                        Console.WriteLine($"File Uploaded to {sessionOptions.HostName}");
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