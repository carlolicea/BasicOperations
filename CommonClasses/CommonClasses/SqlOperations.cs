using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;
using System.Text;


namespace CommonClasses
{
    public static class SqlOperations
    {
        //
        //This beginning is to set a connection string
        private static string integratedSecurity = "False";
        private static string userId = "userId"; //change
        private static string password = "password"; //change
        private static string connectTimeout = "3";
        private static string encrypt = "False";
        private static string trustServerCertificate = "True";
        private static string applicationIntent = "ReadWrite";
        private static string multiSubnetFailover = "False";
        private static string dbServer = "Db Server"; //change
        private static string database = "Db Name"; //change
        public static string sqlConnectionString = "Server=" + dbServer +
                                ";Database=" + database +
                                ";Integrated Security=" + integratedSecurity +
                                ";User Id=" + userId +
                                ";Password=" + password +
                                ";Connect Timeout=" + connectTimeout +
                                ";Encrypt=" + encrypt +
                                ";TrustServerCertificate=" + trustServerCertificate +
                                ";ApplicationIntent=" + applicationIntent +
                                ";MultiSubnetFailover=" + multiSubnetFailover;

        //
        //Open a SQL connection to the server
        public static SqlConnection SqlOpenConnection(string connectionString)
        {
            try
            {
                //Basic sql connection opening
                SqlConnection sqlConnection = new SqlConnection(connectionString);
                sqlConnection.Open();
                return sqlConnection;
            }
            catch
            {
                return null;
            }
        }

        //
        //Get SQL database table names
        public static List<string> SqlGetTableNames(SqlConnection sqlConnection)
        {
            try
            {
                List<string> existingDbTables = new List<string>();

                //Collect the tables with the SqlConnection
                DataTable dt = sqlConnection.GetSchema("Tables");
                foreach (DataRow row in dt.Rows)
                {
                    //The table name will be in column index 2
                    existingDbTables.Add((string)row[2]);
                }
                return existingDbTables;
            }
            catch
            {
                return null;
            }
        }

        //
        //Write a DataTable to SQL database given a table name to use, the SqlConnection, a DataTable with data, and a bool for dropping duplicate tables
        public static void SqlWriteDataTable(string tableName, SqlConnection sqlConnection, DataTable dataTable, bool dropTable)
        {
            //Check if all of the column data is either empty (null) or "NA". If so, remove the column from the incoming data
            foreach (var column in dataTable.Columns.Cast<DataColumn>().ToArray())
            {
                if (dataTable.AsEnumerable().All(dataRow => dataRow.IsNull(column) || dataRow[column].ToString() == "NA"))
                {
                    dataTable.Columns.Remove(column);
                }
            }

            //Check if the count of columns from the incoming DataTable exceeeds 1000. If so, return with the message and exit the operation.
            if (dataTable.Columns.Count > 1000)
            {
                MessageBox.Show("Column count for " + tableName + " exceeds 1,000 column limit.", "Column Count Error");
                return;
            }

            //This string is the Sql Command string if a pre-existing table is to be dropped
            string sqlDropCommand = "IF OBJECT_ID('[dbo].[" + tableName + "]', 'U') IS NOT NULL DROP TABLE [" + tableName + "]";

            //This string is the Sql Command for creating a table given the tableName to use and the DataTable.
            string sqlCreateTableCommand = "CREATE TABLE [" + tableName + "] (";
            for (int i = 0; i < dataTable.Columns.Count; i++) //increment through the columns for adding them to the sqlCreateTableCommand
            {
                sqlCreateTableCommand += "\n [" + dataTable.Columns[i].ColumnName + "] ";
                //Get the DataType of the column and use the switch case to identify what Sql datatype to use.
                string columnType = dataTable.Columns[i].DataType.ToString();
                switch (columnType)
                {
                    case "System.Int32":
                        sqlCreateTableCommand += " int ";
                        break;
                    case "System.Int64":
                        sqlCreateTableCommand += " bigint ";
                        break;
                    case "System.Int16":
                        sqlCreateTableCommand += " smallint ";
                        break;
                    case "System.Byte":
                        sqlCreateTableCommand += " tinyint ";
                        break;
                    case "System.Decimal":
                        sqlCreateTableCommand += " decimal(15,6) "; //15 digits before and 6 digits after the decimal
                        break;
                    case "System.Double":
                        sqlCreateTableCommand += " decimal(15,6) ";
                        break;
                    case "System.DateTime":
                        sqlCreateTableCommand += " datetime ";
                        break;
                    case "System.Guid":
                        sqlCreateTableCommand += " uniqueidentifier ";
                        break;
                    case "System.String":
                        sqlCreateTableCommand += " nvarchar(MAX) ";
                        break;
                    default:
                        sqlCreateTableCommand += string.Format(" nvarchar(MAX) ");
                        break;
                }
                if (dataTable.Columns[i].AutoIncrement)
                    sqlCreateTableCommand += " IDENTITY(" + dataTable.Columns[i].AutoIncrementSeed.ToString() + "," + dataTable.Columns[i].AutoIncrementStep.ToString() + ") ";
                if (!dataTable.Columns[i].AllowDBNull)
                    sqlCreateTableCommand += " NOT NULL ";
                sqlCreateTableCommand += ",";
            } 
            //Example returned result of the loop: CREATE TABLE [SampleTable] (\n [column1] int, \n[column2] datetime,
            sqlCreateTableCommand = sqlCreateTableCommand.Substring(0, sqlCreateTableCommand.Length - 1) + "\n)"; //Removes the last comma and adds \n) to close the command string

            //Use the connection to begin making the Sql table
            using (sqlConnection)
            {
                //Before making the table, drop it if needed
                if (dropTable)
                {
                    SqlCommand sqlTableDrop = new SqlCommand(sqlDropCommand, sqlConnection);
                    sqlTableDrop.ExecuteNonQuery();
                }

                //Get the remaining table names
                List<String> existingTableNames = SqlGetTableNames(sqlConnection);
                SqlBulkCopyOptions sqlBulkCopyOptions = SqlBulkCopyOptions.Default;
                if (!existingTableNames.Contains(tableName))
                {
                    //If the table does not already exist, create it.
                    SqlCommand sqlCreateTable = new SqlCommand(sqlCreateTableCommand, sqlConnection);
                    sqlCreateTable.ExecuteNonQuery();

                    //Next, bulk copy the DataTable into the Sql Table
                    try
                    {
                        //Start the bulk copy
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlConnection, sqlBulkCopyOptions, null))
                        {
                            //Set the destination table created above if it didn't exist
                            sqlBulkCopy.DestinationTableName = "[" + tableName + "]";
                            foreach (var column in dataTable.Columns)
                            {
                                //Cycle through the columns and set the source and destination column map
                                sqlBulkCopy.ColumnMappings.Add(column.ToString(), column.ToString());
                            }
                            //After the columns are mapped, write the data from the DataTable
                            sqlBulkCopy.WriteToServer(dataTable);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Table already exists");
                    }
                }
                else //The table aready exists and needs updated
                {
                    List<string> sqlColumnNamesList = new List<string>();
                    List<string> dataTableColumnNamesList = new List<string>();

                    //Get the columns from the sql table
                    DataTable sqlTableColumns = sqlConnection.GetSchema("Columns", new[] { sqlConnection.DataSource, null, tableName });
                    foreach (DataRow sqlColumnRow in sqlTableColumns.Rows)
                    {
                        //Add the value from the first index in the column, the column name, to the list of column names already in the sql table
                        sqlColumnNamesList.Add(sqlColumnRow[0].ToString());
                    }

                    //Next, get the datatype from the incoming DataTable
                    string sqlDataType;
                    foreach (DataColumn dtColumn in dataTable.Columns)
                    {
                        string columnType = dtColumn.DataType.ToString();
                        switch (columnType)
                        {
                            case "System.Int32":
                                sqlDataType = " int ";
                                break;
                            case "System.Int64":
                                sqlDataType = " bigint ";
                                break;
                            case "System.Int16":
                                sqlDataType = " smallint ";
                                break;
                            case "System.Byte":
                                sqlDataType = " tinyint";
                                break;
                            case "System.Decimal":
                                sqlDataType = " decimal(15,6) ";
                                break;
                            case "System.Double":
                                sqlDataType = " decimal(15,6)";
                                break;
                            case "System.DateTime":
                                sqlDataType = " datetime ";
                                break;
                            case "System.Guid":
                                sqlDataType = " uniqueidentifier ";
                                break;
                            case "System.String":
                            default:
                                sqlDataType = string.Format(" nvarchar(255) ");
                                break;
                        }

                        //If the sqlColumnNamesList does not contain a column with same name as a column in the incoming DataTable...
                        if (!sqlColumnNamesList.Contains(dtColumn.ColumnName))
                        {
                            //Create a new SqlCommand to alter the table to add a new column of the incoming Sql datatype from the DataTable
                            SqlCommand sqlAlterTableCommand = new SqlCommand("ALTER TABLE " + tableName + " ADD " + dtColumn.ColumnName + " " + sqlDataType, sqlConnection);
                            try
                            {
                                sqlAlterTableCommand.ExecuteNonQuery();
                            }
                            catch (Exception e)
                            {
                                MessageBox.Show(e.ToString());
                                continue;
                            }
                        }
                    }

                    //After all new columns are added, bulk copy the data over
                    try
                    {                        
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlConnection, sqlBulkCopyOptions, null))
                        {
                            sqlBulkCopy.DestinationTableName = "[" + tableName + "]";
                            sqlBulkCopy.WriteToServer(dataTable);
                        }
                    }
                    catch (SqlException appendException)
                    {
                        MessageBox.Show(appendException.ToString());
                    }
                }

                //Close the SqlConnection
                sqlConnection.Close();
            }
        }

        //
        //Close a SQL connection to the server
        public static bool SqlCloseConnection(SqlConnection sqlConnection)
        {
            try
            {
                sqlConnection.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
