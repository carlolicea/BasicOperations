using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using RVTDocument = Autodesk.Revit.DB.Document;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using Excel = Microsoft.Office.Interop.Excel;

namespace CommonClasses
{
    public class StringOperations
    {
        //
        //This method will clean up strings to be able to be used as file paths by removing
        //everything that is not acceptable in a path string
        public static string String_CleanFilePath(string originalName)
        {
            try
            {
                //Replace using the regular expression looking for anything that is not
                //(^) a part of the set ([]). The set includes:
                // words (\w)
                // escape characters (\.)
                // @ signs (@)
                // dashes (-)
                // whitespace (\s)
                return Regex.Replace(originalName, @"[^\s\w\.@-]", "", RegexOptions.None, TimeSpan.FromSeconds(1.5));
            }
            catch (RegexMatchTimeoutException)
            {
                return String.Empty;
            }
        }

        //
        //This builds a CSV string from a DataTable
        public static string String_BuildCsvFromDataTable(DataTable dt)
        {
            StringBuilder output = new StringBuilder();
            //The first row of text needs to have the column names separated by commas
            foreach (DataColumn column in dt.Columns)
            {
                //Each column name is encased in quotation marks in case the column name has commas
                var item = column.ColumnName;
                output.AppendFormat(string.Concat("\"", item.ToString(), "\"", ","));
            }
            output.AppendLine();

            //The rows are then evaluated the same way
            foreach (DataRow row in dt.Rows)
            {
                foreach (DataColumn col in dt.Columns)
                {
                    var item = row[col];
                    output.AppendFormat(string.Concat("\"", item.ToString(), "\"", ","));
                }
                output.AppendLine();
            }
            return output.ToString();
        }

        //
        //This builds a string from a DataTable's values
        public static string String_BuildStringFromDataTable(DataTable dt)
        {
            StringBuilder output = new StringBuilder();
            foreach (DataRow row in dt.Rows)
            {
                foreach (DataColumn col in dt.Columns)
                {
                    //Each row is evaluated, then the value for each column
                    var item = row[col];
                    output.AppendFormat(string.Concat(item.ToString(), " "));
                }
                //After a row is evaluated, the output is appended to
                output.AppendLine();
            }
            return output.ToString();
        }

        //
        //This builds a string given a list
        public static string String_BuildStringFromList(List<string> list)
        {
            StringBuilder output = new StringBuilder();
            foreach (string item in list)
            {
                output.AppendFormat(string.Concat(item, " "));
                output.AppendLine();
            }
            return output.ToString();
        }
    }

    public class NumberOperations
    {
        //
        //This pads numbers to add four zero digits to the left of the first digit.        
        public static string Number_PadWithZeros(string input)
        {
            //This returns the regular expression after finding the first number, then putting four zeros in front of it.
            return Regex.Replace(input, @"\d+", match => match.Value.PadLeft(4, '0'));
        }
    }

    public class FileOperations
    {
        //
        //Provides a UI for getting a directory.
        public static string File_GetDirectory()
        {
            string directory = "";
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
            //Show the FolderBrowserDialog and get the returned path
            if (folderBrowserDialog.SelectedPath.ToString() != "")
            {
                directory = folderBrowserDialog.SelectedPath;
            }
            return directory;
        }

        //
        //This prompts the user to select any file
        public static string File_GetFile()
        {
            string file = "";
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.ShowDialog();
            if (fileDialog.FileName.ToString() != "")
            {
                file = fileDialog.FileName;
            }
            return file;
        }

        //
        //This prompts the user to select an Excel file
        public static string File_GetExcelFile()
        {
            string file = "";
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                //The filter is set to only show XLSX files in the dialog
                Filter = "Excel Spreadsheet (*.xlsx)|*xlsx"
            };

            fileDialog.ShowDialog();
            if (fileDialog.FileName.ToString() != "")
            {
                file = fileDialog.FileName;
            }
            return file;
        }

        //
        //Get a Revit Family File
        public static string File_GetRevitFamilyFile()
        {
            string file = "";
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = "RVT Family (*.rfa)|*rfa"
            };
            fileDialog.ShowDialog();
            if (fileDialog.FileName.ToString() != "")
            {
                file = fileDialog.FileName;
            }
            return file;
        }

        //
        //This gets a Revit project file
        public static string File_GetRevitProjectFile()
        {
            string file = "";
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = "RVT Project (*.rvt)|*rvt"
            };
            fileDialog.ShowDialog();
            if (fileDialog.FileName.ToString() != "")
            {
                file = fileDialog.FileName;
            }
            return file;
        }

        //
        //This prompts the user to select a CSV file
        public static string File_GetCsvFile()
        {
            string file = "";
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                //The filter is set to only show CSV files in the dialog
                Filter = "CSV File (*.csv)|*csv"
            };
            fileDialog.ShowDialog();
            if (fileDialog.FileName.ToString() != "")
            {
                file = fileDialog.FileName;
            }
            return file;
        }

        //
        //This writes a CSV string to a file
        public static void File_WriteCsvFromDataTable(DataTable dt, string exportName, string exportDirectory)
        {
            string exportPath = exportDirectory + @"\" + exportName + ".csv";
            if (File.Exists(exportPath))
            {
                File.Delete(exportPath);
            }
            string output = StringOperations.String_BuildCsvFromDataTable(dt);
            File.WriteAllText(exportPath, output);
        }

        //
        //This is a method to delete files in a list without having to define it each time.
        public static void File_DeleteFiles(List<string> filePaths)
        {
            foreach (string path in filePaths)
            {
                File.Delete(path);
            }
        }

        //
        //This gets the file size in MB
        public static double File_GetFileSizeMB(string filePath)
        {
            FileInfo familyFileInfo = new FileInfo(filePath);
            double fileSize = ((double)(familyFileInfo.Length) / 1000000.00);
            return fileSize;
        }
    }

    public class UIOperations
    {
        //
        //These are some colors to use for defaults in a color picker dialog
        public static System.Drawing.Color Teal = System.Drawing.Color.FromArgb(150, 208, 202);
        public static System.Drawing.Color DarkPink = System.Drawing.Color.FromArgb(233, 164, 195);
        public static System.Drawing.Color LightPink = System.Drawing.Color.FromArgb(241, 219, 235);
        public static System.Drawing.Color ClayRed = System.Drawing.Color.FromArgb(224, 155, 144);
        public static System.Drawing.Color BrownTan = System.Drawing.Color.FromArgb(198, 166, 140);
        public static System.Drawing.Color Yellow = System.Drawing.Color.FromArgb(247, 234, 136);
        public static System.Drawing.Color Peach = System.Drawing.Color.FromArgb(245, 221, 195);
        public static System.Drawing.Color CornflowerBlue = System.Drawing.Color.FromArgb(185, 218, 243);
        public static System.Drawing.Color PowderBlue = System.Drawing.Color.FromArgb(217, 234, 236);
        public static System.Drawing.Color DarkMoss = System.Drawing.Color.FromArgb(172, 185, 147);
        public static System.Drawing.Color LightMoss = System.Drawing.Color.FromArgb(219, 234, 184);
        public static System.Drawing.Color Custard = System.Drawing.Color.FromArgb(246, 248, 233);
        public static System.Drawing.Color Slate = System.Drawing.Color.FromArgb(159, 172, 170);
        public static System.Drawing.Color Gray = System.Drawing.Color.FromArgb(220, 221, 222);
        public static System.Drawing.Color LightGray = System.Drawing.Color.FromArgb(245, 246, 246);
        public static System.Drawing.Color White = System.Drawing.Color.FromArgb(255, 255, 255);

        //
        //This creates a custom ColorDialog with default colors
        public static ColorDialog CustomColorDialog()
        {
            int[] customColors = new int[]
            {
                ColorToInt(Teal),
                ColorToInt(DarkPink),
                ColorToInt(LightPink),
                ColorToInt(ClayRed),
                ColorToInt(BrownTan),
                ColorToInt(Yellow),
                ColorToInt(Peach),
                ColorToInt(CornflowerBlue),
                ColorToInt(PowderBlue),
                ColorToInt(DarkMoss),
                ColorToInt(LightMoss),
                ColorToInt(Custard),
                ColorToInt(Slate),
                ColorToInt(Gray),
                ColorToInt(LightGray),
                ColorToInt(White)
            };
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.AllowFullOpen = true;
            colorDialog.CustomColors = customColors;
            return colorDialog;
        }

        //
        //This is required to convert the RGB color to an Int
        private static int ColorToInt(System.Drawing.Color color)
        {
            byte[] result = new byte[4];
            result[0] = color.R;
            result[1] = color.G;
            result[2] = color.B;
            result[3] = 0;
            return BitConverter.ToInt32(result, 0);
        }
    }

    public class DataOperations
    {
        //
        //This simply binds a DataTable to a DataGridView
        public static void Data_BindTableToGridView(DataGridView dgv, DataTable dt)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dt;
            dgv.DataSource = bs;
        }

        //
        //This resets the data in a DataGridView
        public static void Data_ResetGridView(DataGridView dataGridView)
        {
            //Stop all edits to the DGV
            dataGridView.CancelEdit();
            //Clear the columns
            dataGridView.Columns.Clear();
            //Then clear the data source
            dataGridView.DataSource = null;
        }
    }

    public class SqlOperations
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
        public static SqlConnection Sql_OpenConnection(string connectionString)
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
        public static List<string> Sql_GetTableNames(SqlConnection sqlConnection)
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
        public static void Sql_WriteDataTable(string tableName, SqlConnection sqlConnection, DataTable dataTable, bool dropTable)
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
                List<String> existingTableNames = Sql_GetTableNames(sqlConnection);
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
        public static bool Sql_CloseConnection(SqlConnection sqlConnection)
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

    public class RevitOperations
    {
        #region UIOperations
        //
        //Causes the UI to invoke selection of an element
        public static ElementId Revit_SelectElement(UIApplication uiApp)
        {
            ElementId elementId = null;
            Selection selection = uiApp.ActiveUIDocument.Selection;
            Reference elemReference = selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
            if (elemReference != null)
            {
                elementId = elemReference.ElementId;
            }
            return elementId;
        }
        //
        //Invokes selection of rooms or room elements to get room elements
        public static List<Room> Revit_SelectRoomElements(UIApplication uiApp)
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            List<Element> selectedElements = new List<Element>();
            IList<Reference> elemReferences = new List<Reference>();

            try
            {
                elemReferences = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element).Distinct().ToList();
                foreach (Reference selectedReference in elemReferences)
                {
                    ElementId referenceId = selectedReference.ElementId;
                    Element referenceElement = uidoc.Document.GetElement(referenceId);
                    selectedElements.Add(referenceElement);
                }

                Dictionary<int, Room> roomIdsDict = new Dictionary<int, Room>();
                foreach (Element element in selectedElements)
                {
                    if (element.GetType().ToString() == "Autodesk.Revit.DB.Architecture.Room")
                    {
                        Room room = element as Room;
                        if (!roomIdsDict.Keys.Contains(room.Id.IntegerValue))
                        {
                            roomIdsDict[room.Id.IntegerValue] = room;
                        }
                    }
                    else if (element.GetType().ToString() == "Autodesk.Revit.DB.Architecture.RoomTag")
                    {
                        RoomTag tag = element as RoomTag;
                        Room tagRoom = tag.Room;
                        if (!roomIdsDict.Keys.Contains(tag.Room.Id.IntegerValue))
                        {
                            roomIdsDict[tag.Room.Id.IntegerValue] = tagRoom;
                        }
                    }
                    else
                    {
                        continue;
                    }
                }

                List<Room> selectedRoomElements = new List<Room>();
                foreach (int key in roomIdsDict.Keys)
                {
                    selectedRoomElements.Add(roomIdsDict[key]);
                }
                return selectedRoomElements;
            }
            catch
            {
                return null;
            }
        }
        //
        //This invokes selection of elements and filters curve elements
        public static List<Curve> Revit_SelectLineElements(UIApplication uiApp)
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            List<Curve> selectedElements = new List<Curve>();
            ISelectionFilter selectionFilter = new CurveSelectionFilter();
            List<Curve> curveList = uidoc.Selection.PickElementsByRectangle(selectionFilter, "Select Line Elements") as List<Curve>;
            return curveList;
        }
        //
        //Provides a selection filter for only selecting curve elements
        private class CurveSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element element)
            {
                if (element.Category.Name == "Curve")
                {
                    return true;
                }
                return false;
            }
            public bool AllowReference(Reference reference, XYZ point)
            {
                return false;
            }
        }
        //
        //This is to handle the load options
        private class FamilyLoadOptions : IFamilyLoadOptions
        {
            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                return true;
            }
            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Family;
                overwriteParameterValues = true;
                return true;
            }
        }
        #endregion UIOperations


        #region FailureHandlingOperations
        //
        //This is to handle how to handle duplicate types warning
        private class DuplicateTypesHandler : IDuplicateTypeNamesHandler
        {
            public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs e)
            {
                return DuplicateTypeAction.UseDestinationTypes;
            }
        }
        //
        //This is a means to handle failures as they occur
        private class FailuresProcessor
        {
            //
            //This is the default error/warning message handler for when they occur
            public static void OnFailuresProcessing(object sender, FailuresProcessingEventArgs e)
            {
                //Get the FailuresAccessor
                FailuresAccessor failuresAccessor = e.GetFailuresAccessor();
                //Get all messages in the FailuresAccessor
                IList<FailureMessageAccessor> fmas = failuresAccessor.GetFailureMessages();
                //If there are no failures, just continue
                if (fmas.Count == 0)
                {
                    e.SetProcessingResult(FailureProcessingResult.Continue);
                }
                //Otherwwise, evaluate the severity of the failure messages
                else
                {
                    //Cycle through each failure message
                    foreach (FailureMessageAccessor fma in fmas)
                    {
                        try
                        {
                            failuresAccessor.DeleteWarning(fma);
                        }
                        catch
                        {
                            failuresAccessor.ResolveFailure(fma);
                        }
                    }
                }
            }
        }
        //
        //This is a means to handle failures before they post
        private class FailuresPreprocessor : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor fa)
            {
                IList<FailureMessageAccessor> fmas = fa.GetFailureMessages();
                //Cycle through each failure message
                foreach (FailureMessageAccessor fma in fmas)
                {
                    try
                    {
                        fa.DeleteWarning(fma);
                    }
                    catch
                    {
                        fa.ResolveFailure(fma);
                    }
                }
                return FailureProcessingResult.Continue;
            }
        }
        #endregion FailureHandlingOperations


        #region DocumentOperations
        //
        //This gets the major version of the Revit file
        public static string Revit_GetVersion(string filePath)
        {
            if (filePath != null && filePath != "")
            {
                try
                {
                    BasicFileInfo rvtInfo = BasicFileInfo.Extract(filePath);
                    string rvtVersion = rvtInfo.Format.ToString();
                    return rvtVersion;
                }
                catch
                {
                    try
                    {
                        BasicFileInfo rvtInfo = BasicFileInfo.Extract(filePath);
                        string rvtVersion = rvtInfo.Format.ToString();
                        return rvtVersion;
                    }
                    catch
                    {
                        return string.Empty;
                    }
                }
            }
            else { return string.Empty; }
        }
        //
        //This converts the Revit version to a number
        public static int Revit_GetVersionNumber(string filePath)
        {
            int rvtNumber = 0;
            try
            {
                string rvtVersion = Revit_GetVersion(filePath);
                rvtNumber = Convert.ToInt32(rvtVersion.Substring(rvtVersion.Length - 4));
                return rvtNumber;

            }
            catch { return rvtNumber; }
        }
        //
        //Opens a Revit file
        public static RVTDocument Revit_OpenRevitFile(UIApplication uiApp, string filePath)
        {
            RVTDocument doc = null;
            string fileExtension = Path.GetExtension(filePath);
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            if (fileExtension.ToLower() == ".rvt" || fileExtension.ToLower() == ".rfa")
            {
                try
                {
                    doc = uiApp.Application.OpenDocumentFile(filePath);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
            }
            else
            {
                MessageBox.Show(string.Format("{0} is not a Revit file", fileName));
            }
            return doc;
        }
        //
        //This sets all links to overlay
        public static void Revit_SetLinksToOverlay(RVTDocument doc)
        {
            var linkTypes = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).ToElements();

            Transaction t = new Transaction(doc, "SetLinksToOverlay");
            t.Start();
            try
            {
                foreach (Element elem in linkTypes)
                {
                    RevitLinkType linkType = elem as RevitLinkType;
                    linkType.AttachmentType = AttachmentType.Overlay;
                }
                t.Commit();
            }
            catch
            {
                t.RollBack();
            }
        }
        //
        //Reloads links after they were unloaded
        public static TransmissionData Revit_ReloadLinks(string filePath)
        {

            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
            try
            {
                TransmissionData transmissionData = TransmissionData.ReadTransmissionData(modelPath);
                if (transmissionData != null)
                {
                    ICollection<ElementId> externalFileReferences = transmissionData.GetAllExternalFileReferenceIds();
                    foreach (ElementId elemId in externalFileReferences)
                    {
                        ExternalFileReference exRef = transmissionData.GetLastSavedReferenceData(elemId);
                        if (exRef.ExternalFileReferenceType == ExternalFileReferenceType.RevitLink)
                        {
                            transmissionData.SetDesiredReferenceData(elemId, modelPath, PathType.Absolute, true);
                        }
                    }
                }
                transmissionData.IsTransmitted = false;
                TransmissionData.WriteTransmissionData(modelPath, transmissionData);
                return transmissionData;
            }
            catch
            {
                return null;
            }
        }
        //
        //This unloads all links without opening the document
        public static TransmissionData Revit_UnloadLinks(string filePath)
        {
            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
            try
            {
                TransmissionData transmissionData = TransmissionData.ReadTransmissionData(modelPath);
                if (transmissionData != null)
                {
                    ICollection<ElementId> externalFileReferences = transmissionData.GetAllExternalFileReferenceIds();
                    foreach (ElementId elemId in externalFileReferences)
                    {
                        ExternalFileReference exRef = transmissionData.GetLastSavedReferenceData(elemId);
                        if (exRef.ExternalFileReferenceType == ExternalFileReferenceType.RevitLink)
                        {
                            transmissionData.SetDesiredReferenceData(elemId, exRef.GetPath(), PathType.Absolute, false);
                        }
                    }
                }
                transmissionData.IsTransmitted = true;
                TransmissionData.WriteTransmissionData(modelPath, transmissionData);
                return transmissionData;
            }
            catch
            {
                return null;
            }
        }
        //
        //Checks to determine if a Revit file should be allowed to be upgraded
        public static bool Revit_AllowUpgrade(UIApplication uiApp, string filePath, bool allowEqualVersion)
        {
            bool result;
            int fileRevitNumber = Revit_GetVersionNumber(filePath);
            if (allowEqualVersion)
            {
                if (fileRevitNumber <= Convert.ToInt32(uiApp.Application.VersionNumber))
                {
                    result = true;
                }
                else
                {
                    result = false;
                }
                return result;
            }
            else
            {
                if (fileRevitNumber < Convert.ToInt32(uiApp.Application.VersionNumber))
                {
                    result = true;
                }
                else
                {
                    result = false;
                }
                return result;
            }
        }
        //
        //This sets the project upgrade name if it has a Revit version in it
        public static string Revit_SetUpgradedProjectName(UIApplication uiApp, string originalFilePath)
        {

            string upgradeFileName = String.Empty;
            string[] lettersToCheck = new string[] { "A", "a", "V", "v", "R", "r", "S", "s", "M", "m", "E", "e", "P", "p" };
            List<string> tags = new List<string>();
            string originalFileName = Path.GetFileNameWithoutExtension(originalFilePath);
            string changingFileName = Path.GetFileNameWithoutExtension(originalFilePath);

            string fileRevitVersion = Convert.ToString(Revit_GetVersionNumber(originalFilePath));
            string appRevitVersion = uiApp.Application.VersionNumber;

            string fileRevitDigits = fileRevitVersion.Substring(fileRevitVersion.Length - 2);
            string appRevitDigits = appRevitVersion.Substring(appRevitVersion.Length - 2);

            bool skip = false;

            foreach (string x in lettersToCheck)
            {
                tags.Add(string.Join("", x, fileRevitDigits));
            }

            foreach (string tag in tags)
            {
                string replacementTag = tag.Replace(fileRevitDigits, appRevitDigits);
                if (originalFileName.Contains(replacementTag))
                {
                    return originalFileName;
                }
                else
                {
                    changingFileName = changingFileName.Replace(tag, replacementTag);
                }
            }

            if (skip == true)
            {
                upgradeFileName = originalFileName;
            }
            else if (changingFileName == originalFileName && skip == false)
            {
                upgradeFileName = string.Join("", originalFileName, "-A", appRevitDigits);
            }
            else
            {
                upgradeFileName = changingFileName;
            }

            return upgradeFileName;
        }
        //
        //This saves a Revit file back to the original file
        public static bool Revit_SaveFile(RVTDocument doc, bool makeCentral, bool close)
        {
            bool result = false;
            TransactWithCentralOptions TWCOptions = new TransactWithCentralOptions();
            RelinquishOptions relinquishOptions = new RelinquishOptions(true);
            SynchronizeWithCentralOptions SWCOptions = new SynchronizeWithCentralOptions();
            SWCOptions.Compact = true;
            SWCOptions.SetRelinquishOptions(relinquishOptions);
            WorksharingSaveAsOptions worksharingSaveOptions = new WorksharingSaveAsOptions();
            worksharingSaveOptions.SaveAsCentral = true;
            SaveOptions saveOptions = new SaveOptions();
            saveOptions.Compact = true;
            SaveAsOptions saveAsOptions = new SaveAsOptions();
            saveAsOptions.Compact = true;
            saveAsOptions.MaximumBackups = 3;
            saveAsOptions.OverwriteExistingFile = true;

            try
            {
                if (doc.IsFamilyDocument)
                {
                    try
                    {
                        doc.SaveAs(doc.PathName, saveAsOptions);
                        result = true;
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                    }
                }
                else
                {
                    try
                    {
                        Revit_SetLinksToOverlay(doc);
                        if (doc.IsWorkshared)
                        {

                            saveAsOptions.SetWorksharingOptions(worksharingSaveOptions);
                            doc.Save(saveOptions);
                            doc.SynchronizeWithCentral(TWCOptions, SWCOptions);
                        }
                        else if (makeCentral)
                        {

                            saveAsOptions.SetWorksharingOptions(worksharingSaveOptions);
                            doc.EnableWorksharing("Shared Levels and Grids", "Workset1");
                            doc.Save(saveOptions);
                            doc.SynchronizeWithCentral(TWCOptions, SWCOptions);
                        }
                        else
                        {
                            doc.Save(saveOptions);
                        }
                        result = true;
                    }
                    catch (Exception e) { MessageBox.Show(e.Message); doc.Close(); }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                if (close == true)
                {
                    doc.Close(false);
                }
            }
            return result;
        }
        //
        //This saves a Revit file as a new file
        public static bool Revit_SaveAsFile(RVTDocument doc, string saveLocation, bool makeCentral, bool close)
        {
            bool result = false;
            TransactWithCentralOptions TWCOptions = new TransactWithCentralOptions();
            RelinquishOptions relinquishOptions = new RelinquishOptions(true);
            SynchronizeWithCentralOptions SWCOptions = new SynchronizeWithCentralOptions();
            SWCOptions.Compact = true;
            SWCOptions.SetRelinquishOptions(relinquishOptions);
            WorksharingSaveAsOptions worksharingSaveOptions = new WorksharingSaveAsOptions();
            worksharingSaveOptions.SaveAsCentral = true;
            SaveAsOptions saveAsOptions = new SaveAsOptions();
            saveAsOptions.Compact = true;
            saveAsOptions.MaximumBackups = 3;
            saveAsOptions.OverwriteExistingFile = true;

            try
            {
                if (doc.IsFamilyDocument)
                {
                    try
                    {
                        doc.SaveAs(saveLocation, saveAsOptions);
                        result = true;
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                    }
                }
                else
                {
                    try
                    {
                        Revit_SetLinksToOverlay(doc);
                        if (doc.IsWorkshared)
                        {
                            saveAsOptions.SetWorksharingOptions(worksharingSaveOptions);
                            doc.SaveAs(saveLocation, saveAsOptions);
                            doc.SynchronizeWithCentral(TWCOptions, SWCOptions);
                        }
                        else if (makeCentral)
                        {
                            saveAsOptions.SetWorksharingOptions(worksharingSaveOptions);
                            doc.EnableWorksharing("Shared Levels and Grids", "Workset1");
                            doc.SaveAs(saveLocation, saveAsOptions);
                            doc.SynchronizeWithCentral(TWCOptions, SWCOptions);
                        }
                        else
                        {
                            doc.SaveAs(saveLocation, saveAsOptions);
                        }
                        result = true;
                    }
                    catch (Exception e) { MessageBox.Show(e.Message); doc.Close(); }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                if (close == true)
                {
                    doc.Close(false);
                }
            }
            return result;
        }
        //
        //This will delete all Revit family backups
        public static void Revit_CleanRfaBackups(List<string> filePaths)
        {
            //Ensure the incoming list of file paths is not empty or null
            if (filePaths.Count != 0 && filePaths != null)
            {
                foreach (string filepath in filePaths)
                {
                    //Use a regular expression to find the .00##.rfa pattern in the file name.
                    Match match = Regex.Match(Path.GetFileName(filepath), @".00\d\d.rfa", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        //If the family is identified as a Revit backup, delete it.
                        try
                        {
                            File.Delete(filepath);
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show(e.ToString());
                        }
                    }
                }
            }
        }
        //
        //This method is used to collect paths for all Revit family backups given a directory
        public static List<string> Revit_GetAllFamilyBackupPaths(string directoryPath, bool searchSubDirectories)
        {
            List<string> filePaths = new List<string>();
            if (directoryPath != "")
            {
                //If all subdirectories shouldbe searched, continue. If only the top level of the supplied directory is
                //to be searched, go to the ELSE operations.
                if (searchSubDirectories)
                {
                    foreach (string filePath in Directory.GetFiles(directoryPath, "*.rfa", SearchOption.AllDirectories))
                    {
                        //Determine if the file name contains the following pattern: .00##.rfa
                        Match match = Regex.Match(Path.GetFileName(filePath), @".00\d\d.rfa", RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            //If the file name is a Revit family backup, add it to the list.
                            filePaths.Add(filePath);
                        }
                    }
                }
                else
                {
                    //The following is the same as the IF operations, but only searches the folder of the directory supplied
                    foreach (string filePath in Directory.GetFiles(directoryPath, "*.rfa", SearchOption.TopDirectoryOnly))
                    {
                        Match match = Regex.Match(Path.GetFileName(filePath), @".00\d\d.rfa", RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            filePaths.Add(filePath);
                        }
                    }
                }
            }
            return filePaths;
        }
        //
        //Gets all Revit RVT files from a directory, passing a last modified date newer than the one supplied. This also takes the log file to record failures
        public static List<string> Revit_GetAllProjectFiles(string directoryPath, DateTime date)
        {
            List<string> files = new List<string>();
            //Get an array of directories given the path
            string[] directories = Directory.GetDirectories(directoryPath);
            //Cyle through the directories
            foreach (string directory in directories)
            {
                //Wrapping this in a TRY/CATCH because some directories may not be accessible
                try
                {
                    //Get the files from the directory and add them to a list
                    List<string> filePaths = Directory.EnumerateFiles(directory, "*.rvt", SearchOption.AllDirectories).ToList();

                    //Cycle through the file paths
                    foreach (string file in filePaths)
                    {
                        FileAttributes fileAttributes = File.GetAttributes(file);
                        //Wrapping this in a TRY/CATCH because some files may not be accessible
                        try
                        {
                            //Change the attributes to normal to disable any Read Only attribute.
                            File.SetAttributes(file, FileAttributes.Normal);
                            //Get the FileInfo from the file
                            FileInfo fileInfo = new FileInfo(file);
                            //If the LastWriteTime is newer than the date supplied to the method, add the file path to the list of files
                            if (fileInfo.LastWriteTime >= date)
                            {
                                files.Add(file);
                            }
                        }
                        catch
                        {
                            continue;
                        }
                        finally
                        {
                            //Reset the file attributes to what they were prior to removing any Read Only attribute
                            File.SetAttributes(file, fileAttributes);
                        }
                    }
                }
                catch { continue; }
            }
            return files;
        }
        //
        //This gets the BuiltInParameterGroup given the name
        public static BuiltInParameterGroup Revit_GetBuiltInParameterGroupFromString(string groupName)
        {
            BuiltInParameterGroup group = BuiltInParameterGroup.INVALID;
            switch (groupName)
            {
                case "Analysis Results":
                    group = BuiltInParameterGroup.PG_ANALYSIS_RESULTS;
                    break;
                case "Analytical Alignment":
                    group = BuiltInParameterGroup.PG_ANALYTICAL_ALIGNMENT;
                    break;
                case "Analytical Model":
                    group = BuiltInParameterGroup.PG_ANALYTICAL_MODEL;
                    break;
                case "Constraints":
                    group = BuiltInParameterGroup.PG_CONSTRAINTS;
                    break;
                case "Construction":
                    group = BuiltInParameterGroup.PG_CONSTRUCTION;
                    break;
                case "Data":
                    group = BuiltInParameterGroup.PG_DATA;
                    break;
                case "Dimensions":
                    group = BuiltInParameterGroup.PG_GEOMETRY;
                    break;
                case "Division Geometry":
                    group = BuiltInParameterGroup.PG_DIVISION_GEOMETRY;
                    break;
                case "Electrical":
                    group = BuiltInParameterGroup.PG_ELECTRICAL;
                    break;
                case "Electrical - Circuiting":
                    group = BuiltInParameterGroup.PG_ELECTRICAL_CIRCUITING;
                    break;
                case "Electrical - Lighting":
                    group = BuiltInParameterGroup.PG_ELECTRICAL_LIGHTING;
                    break;
                case "Electrical - Loads":
                    group = BuiltInParameterGroup.PG_ELECTRICAL_LOADS;
                    break;
                case "Electrical Engineering":
                    group = BuiltInParameterGroup.PG_AELECTRICAL;
                    break;
                case "Energy Analysis":
                    group = BuiltInParameterGroup.PG_ENERGY_ANALYSIS;
                    break;
                case "Fire Protection":
                    group = BuiltInParameterGroup.PG_FIRE_PROTECTION;
                    break;
                case "Forces":
                    group = BuiltInParameterGroup.PG_FORCES;
                    break;
                case "General":
                    group = BuiltInParameterGroup.PG_GENERAL;
                    break;
                case "Graphics":
                    group = BuiltInParameterGroup.PG_GRAPHICS;
                    break;
                case "Green Building Properties":
                    group = BuiltInParameterGroup.PG_GREEN_BUILDING;
                    break;
                case "Identity Data":
                    group = BuiltInParameterGroup.PG_IDENTITY_DATA;
                    break;
                case "IFC Parameters":
                    group = BuiltInParameterGroup.PG_IFC;
                    break;
                case "Materials and Finishes":
                    group = BuiltInParameterGroup.PG_MATERIALS;
                    break;
                case "Mechanical":
                    group = BuiltInParameterGroup.PG_MECHANICAL;
                    break;
                case "Mechanical - Flow":
                    group = BuiltInParameterGroup.PG_MECHANICAL_AIRFLOW;
                    break;
                case "Mechanical - Loads":
                    group = BuiltInParameterGroup.PG_MECHANICAL_LOADS;
                    break;
                case "Model Properties":
                    group = BuiltInParameterGroup.PG_ADSK_MODEL_PROPERTIES;
                    break;
                case "Moments":
                    group = BuiltInParameterGroup.PG_MOMENTS;
                    break;
                case "Other":
                    group = BuiltInParameterGroup.INVALID;
                    break;
                case "Overall Legend":
                    group = BuiltInParameterGroup.PG_OVERALL_LEGEND;
                    break;
                case "Phasing":
                    group = BuiltInParameterGroup.PG_PHASING;
                    break;
                case "Photometrics":
                    group = BuiltInParameterGroup.PG_LIGHT_PHOTOMETRICS;
                    break;
                case "Plumbing":
                    group = BuiltInParameterGroup.PG_PLUMBING;
                    break;
                case "Primary End":
                    group = BuiltInParameterGroup.PG_PRIMARY_END;
                    break;
                case "Rebar Set":
                    group = BuiltInParameterGroup.PG_REBAR_ARRAY;
                    break;
                case "Releases / Member Forces":
                    group = BuiltInParameterGroup.PG_RELEASES_MEMBER_FORCES;
                    break;
                case "Secondary End":
                    group = BuiltInParameterGroup.PG_SECONDARY_END;
                    break;
                case "Segments and Fittings":
                    group = BuiltInParameterGroup.PG_SEGMENTS_FITTINGS;
                    break;
                case "Slab Shape Edit":
                    group = BuiltInParameterGroup.PG_SLAB_SHAPE_EDIT;
                    break;
                case "Structural":
                    group = BuiltInParameterGroup.PG_STRUCTURAL;
                    break;
                case "Structural Analysis":
                    group = BuiltInParameterGroup.PG_STRUCTURAL_ANALYSIS;
                    break;
                case "Text":
                    group = BuiltInParameterGroup.PG_TEXT;
                    break;
                case "Title Text":
                    group = BuiltInParameterGroup.PG_TITLE;
                    break;
                case "Visibility":
                    group = BuiltInParameterGroup.PG_VISIBILITY;
                    break;
                default:
                    MessageBox.Show(String.Format("Could not get the BuiltInParameterGroup {0}", groupName));
                    break;
            }
            return group;
        }
        //
        //This gets the name of a BuiltInParameterGroup
        public static string Revit_GetNameOfBuiltInParameterGroup(BuiltInParameterGroup paramGroup)
        {
            string group = "Other";
            switch (paramGroup)
            {
                case BuiltInParameterGroup.PG_ANALYSIS_RESULTS:
                    group = "Analysis Results";
                    break;
                case BuiltInParameterGroup.PG_ANALYTICAL_ALIGNMENT:
                    group = "Analytical Alignment";
                    break;
                case BuiltInParameterGroup.PG_ANALYTICAL_MODEL:
                    group = "Analytical Model";
                    break;
                case BuiltInParameterGroup.PG_CONSTRAINTS:
                    group = "Constraints";
                    break;
                case BuiltInParameterGroup.PG_CONSTRUCTION:
                    group = "Construction";
                    break;
                case BuiltInParameterGroup.PG_DATA:
                    group = "Data";
                    break;
                case BuiltInParameterGroup.PG_GEOMETRY:
                    group = "Dimensions";
                    break;
                case BuiltInParameterGroup.PG_DIVISION_GEOMETRY:
                    group = "Division Geometry";
                    break;
                case BuiltInParameterGroup.PG_ELECTRICAL:
                    group = "Electrical";
                    break;
                case BuiltInParameterGroup.PG_ELECTRICAL_CIRCUITING:
                    group = "Electrical - Circuiting";
                    break;
                case BuiltInParameterGroup.PG_ELECTRICAL_LIGHTING:
                    group = "Electrical - Lighting";
                    break;
                case BuiltInParameterGroup.PG_ELECTRICAL_LOADS:
                    group = "Electrical - Loads";
                    break;
                case BuiltInParameterGroup.PG_AELECTRICAL:
                    group = "Electrical Engineering";
                    break;
                case BuiltInParameterGroup.PG_ENERGY_ANALYSIS:
                    group = "Energy Analysis";
                    break;
                case BuiltInParameterGroup.PG_FIRE_PROTECTION:
                    group = "Fire Protection";
                    break;
                case BuiltInParameterGroup.PG_FORCES:
                    group = "Forces";
                    break;
                case BuiltInParameterGroup.PG_GENERAL:
                    group = "General";
                    break;
                case BuiltInParameterGroup.PG_GRAPHICS:
                    group = "Graphics";
                    break;
                case BuiltInParameterGroup.PG_GREEN_BUILDING:
                    group = "Green Building Properties";
                    break;
                case BuiltInParameterGroup.PG_IDENTITY_DATA:
                    group = "Identity Data";
                    break;
                case BuiltInParameterGroup.PG_IFC:
                    group = "IFC Parameters";
                    break;
                case BuiltInParameterGroup.PG_MATERIALS:
                    group = "Materials and Finishes";
                    break;
                case BuiltInParameterGroup.PG_MECHANICAL:
                    group = "Mechanical";
                    break;
                case BuiltInParameterGroup.PG_MECHANICAL_AIRFLOW:
                    group = "Mechanical - Flow";
                    break;
                case BuiltInParameterGroup.PG_MECHANICAL_LOADS:
                    group = "Mechanical - Loads";
                    break;
                case BuiltInParameterGroup.PG_ADSK_MODEL_PROPERTIES:
                    group = "Model Properties";
                    break;
                case BuiltInParameterGroup.PG_MOMENTS:
                    group = "Moments";
                    break;
                case BuiltInParameterGroup.INVALID:
                    group = "Other";
                    break;
                case BuiltInParameterGroup.PG_OVERALL_LEGEND:
                    group = "Overall Legend";
                    break;
                case BuiltInParameterGroup.PG_PHASING:
                    group = "Phasing";
                    break;
                case BuiltInParameterGroup.PG_LIGHT_PHOTOMETRICS:
                    group = "Photometrics";
                    break;
                case BuiltInParameterGroup.PG_PLUMBING:
                    group = "Plumbing";
                    break;
                case BuiltInParameterGroup.PG_PRIMARY_END:
                    group = "Primary End";
                    break;
                case BuiltInParameterGroup.PG_REBAR_ARRAY:
                    group = "Rebar Set";
                    break;
                case BuiltInParameterGroup.PG_RELEASES_MEMBER_FORCES:
                    group = "Releases / Member Forces";
                    break;
                case BuiltInParameterGroup.PG_SECONDARY_END:
                    group = "Secondary End";
                    break;
                case BuiltInParameterGroup.PG_SEGMENTS_FITTINGS:
                    group = "Segments and Fittings";
                    break;
                case BuiltInParameterGroup.PG_SLAB_SHAPE_EDIT:
                    group = "Slab Shape Edit";
                    break;
                case BuiltInParameterGroup.PG_STRUCTURAL:
                    group = "Structural";
                    break;
                case BuiltInParameterGroup.PG_STRUCTURAL_ANALYSIS:
                    group = "Structural Analysis";
                    break;
                case BuiltInParameterGroup.PG_TEXT:
                    group = "Text";
                    break;
                case BuiltInParameterGroup.PG_TITLE:
                    group = "Title Text";
                    break;
                case BuiltInParameterGroup.PG_VISIBILITY:
                    group = "Visibility";
                    break;
                default:
                    break;
            }
            return group;
        }
        //
        //This gets all drafting views in the project
        public static List<ViewDrafting> Revit_GetDocumentDraftingViews(UIApplication uiApp)
        {
            RVTDocument doc = uiApp.ActiveUIDocument.Document;

            List<ViewDrafting> draftingViews = new FilteredElementCollector(doc).OfClass(typeof(ViewDrafting)).WhereElementIsNotElementType().ToElements().Cast<ViewDrafting>().ToList();
            return draftingViews;
        }
        //
        //This gets all floor types in the document
        public static List<FloorType> Revit_DocumentFloorTypes(UIApplication uiApp)
        {
            List<FloorType> floorTypes = new List<FloorType>();
            UIDocument uiDoc = uiApp.ActiveUIDocument;

            var floorTypeCollector = new FilteredElementCollector(uiDoc.Document).OfClass(typeof(FloorType)).WhereElementIsElementType().ToElements();
            foreach (Element elem in floorTypeCollector)
            {
                FloorType floorType = elem as FloorType;
                floorTypes.Add(floorType);
            }
            return floorTypes;
        }
        //
        //This gets all floor type names in the document
        public static List<string> Revit_DocumentFloorTypeNames(UIApplication uiApp)
        {
            List<string> floorTypeNames = new List<string>();
            UIDocument uiDoc = uiApp.ActiveUIDocument;

            var floorTypeCollector = new FilteredElementCollector(uiDoc.Document).OfClass(typeof(FloorType)).WhereElementIsElementType().ToElements();
            foreach (Element elem in floorTypeCollector)
            {
                FloorType floorType = elem as FloorType;
                floorTypeNames.Add(floorType.Name.ToString());
            }
            return floorTypeNames;
        }
        //
        //This gets the line styles in the document
        public static CategoryNameMap DocumentLineStyles(UIApplication uiApp)
        {
            CategoryNameMap lineStyleSubCats = null;
            var lineStyles = uiApp.ActiveUIDocument.Document.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            lineStyleSubCats = lineStyles.SubCategories;
            return lineStyleSubCats;
        }
        //
        //This gets all family types given a family name in the document
        public static List<FamilySymbol> Revit_GetFamilyTypesByFamilyName(UIApplication uiApp, string familyName)
        {
            try
            {
                List<FamilySymbol> familySymbols = new List<FamilySymbol>();
                var familySymbolsCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document).OfClass(typeof(FamilySymbol)).ToList();
                foreach (Element elem in familySymbolsCollector)
                {
                    FamilySymbol symbol = elem as FamilySymbol;
                    if (symbol.FamilyName == familyName)
                    {
                        familySymbols.Add(symbol);
                    }
                }
                return familySymbols;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return null;
            }
        }
        //
        //This gets a family by its family name
        public static Family Revit_FamilyByFamilyName(UIApplication uiApp, string familyName)
        {
            Family family = null;
            try
            {
                family = new FilteredElementCollector(uiApp.ActiveUIDocument.Document).OfClass(typeof(Family)).WhereElementIsNotElementType().Where(f => f.Name == familyName).First() as Family;
                return family;
            }
            catch
            {
                return family;
            }
        }
        #endregion DocumentOperations


        #region FamilyOperations
        //
        //This creates a dictionary of family types indexed by type name
        public static Dictionary<string, FamilyType> Revit_GetFamilyTypeNames(FamilyManager famMan)
        {
            Dictionary<string, FamilyType> types = new Dictionary<string, FamilyType>();
            FamilyTypeSet famTypes = famMan.Types;
            if (famTypes.Size != 0)
            {
                foreach (FamilyType type in famTypes)
                {
                    types.Add(type.Name, type);
                }
                return types;
            }
            else
            {
                return null;
            }

        }
        //
        //This deletes all family types in a Revit family
        public static void Revit_DeleteFamilyTypes(RVTDocument famDoc, FamilyManager famMan)
        {
            int numberOfPreExistingTypes = famMan.Types.Size;
            if (numberOfPreExistingTypes > 1)
            {
                Transaction t1 = new Transaction(famDoc, "DeletePreExistingTypes");
                t1.Start();
                foreach (FamilyType type in famMan.Types)
                {
                    try
                    {
                        famMan.DeleteCurrentType();
                    }
                    catch { break; }
                }
                t1.Commit();
            }
        }
        //
        //This gets the Category of a Revit family
        public static string Revit_GetFamilyCategory(RVTDocument doc)
        {
            string familyCategory = null;
            if (doc.IsFamilyDocument)
            {
                try
                {
                    Family ownerFamily = doc.OwnerFamily;
                    familyCategory = ownerFamily.FamilyCategory.Name.ToString();
                }
                catch
                {
                    string filePath = doc.PathName;
                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    MessageBox.Show(string.Format("Could not get the Revit family category for {0}", fileName));
                }
            }
            return familyCategory;
        }
        //
        //This creates family types given a DataGridView, and saves the family to a new location
        public static void Revit_CreateFamilyTypesFromDgv(UIApplication uiApp, string saveDirectory, DataGridView dgv, string familyFileToUse)
        {
            RVTDocument famDoc = Revit_OpenRevitFile(uiApp, familyFileToUse);

            SaveAsOptions saveAsOptions = new SaveAsOptions();
            saveAsOptions.Compact = true;
            saveAsOptions.MaximumBackups = 1;
            saveAsOptions.OverwriteExistingFile = true;

            FamilyManager famMan = famDoc.FamilyManager;
            Revit_DeleteFamilyTypes(famDoc, famMan);

            string tempFamilyPath = saveDirectory + "\\" + String.Format(famDoc.Title).Replace(".rfa", "") + "_temp.rfa";
            famDoc.SaveAs(tempFamilyPath, saveAsOptions);
            famDoc.Close();

            RVTDocument newFamDoc = Revit_OpenRevitFile(uiApp, tempFamilyPath);
            FamilyManager newFamMan = newFamDoc.FamilyManager;

            FamilyParameterSet parameters = newFamMan.Parameters;
            Dictionary<string, FamilyParameter> famParamDict = new Dictionary<string, FamilyParameter>();
            foreach (FamilyParameter param in parameters)
            {
                famParamDict.Add(param.Definition.Name, param);
            }
            int rowsCount = dgv.Rows.Count;
            int columnsCount = dgv.Columns.Count;

            List<string> familyTypesMade = new List<string>();
            Transaction t2 = new Transaction(newFamDoc, "MakeNewTypes");
            t2.Start();
            for (int i = 0; i < rowsCount; i++)
            {
                string newTypeName = dgv.Rows[i].Cells[0].Value.ToString();
                Dictionary<string, FamilyType> existingTypeNames = Revit_GetFamilyTypeNames(newFamMan);
                if (!existingTypeNames.Keys.Contains(newTypeName))
                {
                    FamilyType newType = newFamMan.NewType(newTypeName);
                    newFamMan.CurrentType = newType;
                    familyTypesMade.Add(newType.Name);
                }
                else
                {
                    newFamMan.CurrentType = existingTypeNames[newTypeName];
                    familyTypesMade.Add(newFamMan.CurrentType.Name);
                }

                for (int j = 1; j < columnsCount; j++)
                {
                    string paramName = dgv.Columns[j].HeaderText;
                    string paramStorageTypeString = dgv.Rows[0].Cells[j].Value.ToString();
                    var paramValue = dgv.Rows[i].Cells[j].Value;
                    if (paramValue.ToString() != "")
                    {
                        FamilyParameter param = famParamDict[paramName];
                        ParameterType paramType = param.Definition.ParameterType;
                        Revit_SetFamilyParameterValue(newFamMan, param, paramType, paramStorageTypeString, paramValue, true);
                    }
                }
            }
            t2.Commit();
            Transaction t3 = new Transaction(newFamDoc, "DeleteOldTypes");
            t3.Start();
            foreach (FamilyType type in newFamMan.Types)
            {
                if (!familyTypesMade.Contains(type.Name))
                {
                    newFamMan.CurrentType = type;
                    newFamMan.DeleteCurrentType();
                }
            }
            t3.Commit();
            string nonTempFamilyPath = saveDirectory + "\\" + String.Format(newFamDoc.Title).Replace("_temp", "");
            newFamDoc.SaveAs(nonTempFamilyPath, saveAsOptions);
            newFamDoc.Close();

            File.Delete(tempFamilyPath);
            File.Delete(tempFamilyPath.Replace(".rfa", ".0001.rfa"));
            File.Delete(nonTempFamilyPath.Replace(".rfa", ".0001.rfa"));
        }
        //
        //This creates family types given a DataGridView, and saves it back to the original file, returning the family document
        public static RVTDocument Revit_CreateFamilyTypesFromDgv(UIApplication uiApp, DataGridView dgv, Family familyToUse)
        {
            RVTDocument famDoc = uiApp.ActiveUIDocument.Document.EditFamily(familyToUse);


            FamilyManager famMan = famDoc.FamilyManager;

            FamilyParameterSet parameters = famMan.Parameters;
            Dictionary<string, FamilyParameter> famParamDict = new Dictionary<string, FamilyParameter>();
            foreach (FamilyParameter param in parameters)
            {
                famParamDict.Add(param.Definition.Name, param);
            }
            int rowsCount = dgv.Rows.Count;
            int columnsCount = dgv.Columns.Count;

            List<string> familyTypesMade = new List<string>();
            Transaction t2 = new Transaction(famDoc, "MakeNewTypes");
            t2.Start();
            for (int i = 0; i < rowsCount; i++)
            {
                string newTypeName = dgv.Rows[i].Cells[0].Value.ToString();
                Dictionary<string, FamilyType> existingTypeNames = Revit_GetFamilyTypeNames(famMan);
                if (!existingTypeNames.Keys.Contains(newTypeName))
                {
                    FamilyType newType = famMan.NewType(newTypeName);
                    famMan.CurrentType = newType;
                    familyTypesMade.Add(newType.Name);
                }
                else
                {
                    famMan.CurrentType = existingTypeNames[newTypeName];
                    familyTypesMade.Add(famMan.CurrentType.Name);
                }

                for (int j = 1; j < columnsCount; j++)
                {
                    string paramName = dgv.Columns[j].HeaderText;
                    string paramStorageTypeString = dgv.Rows[0].Cells[j].Value.ToString();
                    var paramValue = dgv.Rows[i].Cells[j].Value;
                    if (paramValue.ToString() != "")
                    {
                        FamilyParameter param = famParamDict[paramName];
                        ParameterType paramType = param.Definition.ParameterType;
                        if (!param.IsDeterminedByFormula)
                        {
                            Revit_SetFamilyParameterValue(famMan, param, paramType, paramStorageTypeString, paramValue, true);
                        }
                    }
                }
            }
            t2.Commit();
            return famDoc;
        }
        //
        //Sets a parameter value in an open family given a parameter type and possible inch values
        public static void Revit_SetFamilyParameterValue(FamilyManager famMan, FamilyParameter param, ParameterType paramType, string paramStorageTypeString, object paramValue, bool convertInchestoFeet)
        {
            try
            {
                if (paramStorageTypeString == "Integer")
                {
                    famMan.Set(param, Convert.ToInt32(paramValue));
                }
                else if (paramStorageTypeString == "Double")
                {
                    if (paramType.ToString() == "Length" && convertInchestoFeet == true)
                    {
                        famMan.Set(param, Convert.ToDouble(paramValue) / 12d);
                    }
                    else if (paramType.ToString() == "Length" && convertInchestoFeet == false)
                    {
                        famMan.Set(param, Convert.ToDouble(paramValue));
                    }
                    else
                    {
                        famMan.Set(param, Convert.ToDouble(paramValue));
                    }
                }
                else if (paramStorageTypeString == "ElementId")
                {
                    famMan.Set(param, new ElementId(Convert.ToInt32(paramValue)));
                }
                else
                {
                    famMan.Set(param, Convert.ToString(paramValue));
                }
            }
            catch { MessageBox.Show(String.Format("Could not set parameter ({0}) with value ({1}) for type ({2})", param.Definition.Name, paramValue.ToString(), famMan.CurrentType.Name)); }
        }
        //
        //Sets a parameter value through without a given parameter type
        public static void Revit_SetFamilyParameterValue(FamilyManager famMan, FamilyParameter param, object paramValue)
        {
            try
            {
                string paramStorageTypeString = param.StorageType.ToString();
                if (paramStorageTypeString == "Integer")
                {
                    famMan.Set(param, Convert.ToInt32(paramValue));
                }
                else if (paramStorageTypeString == "Double")
                {
                    famMan.Set(param, Convert.ToDouble(paramValue) / 12d);
                }
                else if (paramStorageTypeString == "ElementId")
                {
                    famMan.Set(param, new ElementId(Convert.ToInt32(paramValue)));
                }
                else
                {
                    famMan.Set(param, Convert.ToString(paramValue));
                }
            }
            catch { MessageBox.Show(String.Format("Could not set parameter {0} with value {1}", param.Definition.Name, paramValue.ToString())); }
        }
        //
        //This sets a parameter value given a Revit parameter type
        public static object Revit_SetParameterByParameterType(string typeName, object value)
        {
            object returnValue = null;
            switch (typeName)
            {
                case "Text":
                    returnValue = Convert.ToString(value);
                    break;
                case "Integer":
                    returnValue = Convert.ToInt32(value);
                    break;
                case "Number":
                    returnValue = Convert.ToDouble(value);
                    break;
                case "Length":
                    returnValue = Convert.ToDouble(value);
                    break;
                case "Area":
                    returnValue = Convert.ToDouble(value);
                    break;
                case "Volume":
                    returnValue = Convert.ToDouble(value);
                    break;
                case "Angle":
                    returnValue = Convert.ToDouble(value);
                    break;
                case "Slope":
                    returnValue = Convert.ToDouble(value);
                    break;
                case "Currencey":
                    returnValue = Convert.ToDouble(value);
                    break;
                case "Mass Density":
                    returnValue = Convert.ToDouble(value);
                    break;
                case "URL":
                    returnValue = Convert.ToString(value);
                    break;
                case "Material":
                    returnValue = new ElementId(Convert.ToInt32(value));
                    break;
                case "Image":
                    returnValue = ParameterType.Image;
                    break;
                case "Yes/No":
                    if (Convert.ToBoolean(value) == true)
                    { returnValue = 1; }
                    else if (Convert.ToBoolean(value) == false)
                    { returnValue = 0; }
                    else
                    { returnValue = null; }
                    break;
                case "Multiline Text":
                    returnValue = Convert.ToString(value);
                    break;
                case "<Family Type...>":
                    returnValue = new ElementId(Convert.ToInt32(value));
                    break;
                default:
                    returnValue = ParameterType.Invalid;
                    break;
            }
            return returnValue;
        }
        #endregion FamilyOperations


        #region ElementOperations
        //
        //Deletes parts associated with an element
        public static void Revit_DeleteParts(UIApplication uiApp, RVTDocument doc, ElementId elementId)
        {
            doc.Delete(PartUtils.GetAssociatedPartMaker(uiApp.ActiveUIDocument.Document, elementId).Id);
        }
        #endregion ElementOperations
    }

    public class ExcelOperations
    {
        //
        //This starts Excel
        public static Excel.Application Excel_StartApplication()
        {
            //Per the other methods, open Excel
            Excel.Application excel = new Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;

            return excel;
        }

        //
        //This is to open an Excel document as the active workbook
        public static Excel.Workbook Excel_OpenWorkbook(Excel.Application excel, string filePath)
        {
            //If Excel is running
            if (excel != null)
            {
                try
                {
                    //Open the workbook at the file path and return it
                    Excel.Workbook workbook = excel.Workbooks.Open(filePath);
                    return workbook;
                }
                catch
                {
                    //The file may be open
                    MessageBox.Show("Could not open the Excel file. Please verify it is not currently open.");
                    return null;
                }
            }
            else
            {
                //Excel may have not opened
                MessageBox.Show("Excel is currently not running");
                return null;
            }
        }

        //
        //This is used to save and close Excel
        public static void Excel_CloseExcel(Excel.Application excel, Excel.Workbook workbook, bool save)
        {
            //Save the workbook if needed, close it, then stop Excel
            if (save)
            {
                workbook.Save();
            }

            if (workbook != null)
            {
                workbook.Close();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }

        //
        //This converts a DataGridView data to an Excel spreadsheet given an Excel file
        public static void Excel_DataGridViewToExcel(string filePath, DataGridView dgv)
        {
            Excel.Application excel = Excel_StartApplication();
            Excel.Workbook workbook = Excel_OpenWorkbook(excel, filePath);

            //Make sure Excel and a workbook are open
            if (excel != null && workbook != null)
            {
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                //This is to get the column headers of the DGV
                //Excel's first column index is 1
                for (int c = 1; c <= dgv.Columns.Count; c++)
                {
                    try
                    {
                        //Excel orders cells by [row,column]. Because Excel starts indexes at 1, subtract 1 for the 0 index in the DGV
                        worksheet.Cells[1, c] = dgv.Columns[c - 1].HeaderText;
                    }
                    catch
                    {
                        continue;
                    }
                }

                //This is to get the rows of the DGV
                for (int c = 1; c <= dgv.Columns.Count; c++)
                {
                    try
                    {
                        //The first row in the spreadsheet is the column headers, so start on row 2
                        for (int r = 2; r <= dgv.Rows.Count + 1; r++)
                        {
                            //The cell values from the DGV are set in the Excel spreadsheet
                            worksheet.Cells[r, c] = dgv.Rows[r - 2].Cells[c - 1].Value.ToString();
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
                Excel_CloseExcel(excel, workbook, true);
            }
            else
            {
                return;
            }
        }

        //
        //This provides the same functions as Excel_DataGridViewToExcel, but uses a DataTable. See its comments for explanation
        public static void Excel_DataTableToExcel(string filePath, DataTable dt)
        {
            Excel.Application excel = Excel_StartApplication();
            Excel.Workbook workbook = Excel_OpenWorkbook(excel, filePath);

            if (excel != null && workbook != null)
            {
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                for (int c = 1; c < dt.Columns.Count; c++)
                {
                    try
                    {
                        worksheet.Cells[1, c] = dt.Columns[c - 1].ColumnName;
                    }
                    catch
                    {
                        continue;
                    }
                }

                for (int c = 1; c <= dt.Columns.Count; c++)
                {
                    try
                    {
                        for (int r = 2; r <= dt.Rows.Count + 1; r++)
                        {
                            worksheet.Cells[r, c] = dt.Rows[r - 2][c - 1].ToString();
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
                Excel_CloseExcel(excel, workbook, true);
            }
            else
            {
                return;
            }
        }

        //
        //This converts an Excel spreadsheet table to a DataTable
        public static DataTable Excel_ExcelToDataTable(string filePath, bool hasHeader)
        {
            DataTable dt = new DataTable();

            //Per the other methods, open Excel
            Excel.Application excel = Excel_StartApplication();
            Excel.Workbook workbook = Excel_OpenWorkbook(excel, filePath);

            if (excel != null && workbook != null)
            {
                Excel.Worksheet worksheet = workbook.ActiveSheet;
                Excel.Range last = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

                //Get the range of rows and columns used in the spreadsheet
                int lastUsedRow = last.Row;
                int lastUsedColumn = last.Column;

                //if the spreadsheet has a header...
                if (hasHeader)
                {
                    for (int i = 1; i <= lastUsedColumn; i++)
                    {
                        //... use the cell values to define the columns in the DataTable
                        if (!dt.Columns.Contains(worksheet.Cells[1, i].Value2.ToString()) && worksheet.Cells[1, i].Value2.ToString() != "")
                        {
                            dt.Columns.Add(worksheet.Cells[1, i].Value2.ToString());
                        }
                        else
                        { continue; }
                    }
                }
                //... otherwise just add columns with default Column# headers
                else
                {
                    for (int i = 1; i <= lastUsedColumn; i++)
                    {
                        dt.Columns.Add("Column" + Convert.ToString(i));
                    }
                }

                //Count the number of columns created in the DataTable
                int columnCount = dt.Columns.Count;

                //Start adding new rows and filling out the data for the DataTable
                for (int j = 2; j <= lastUsedRow; j++)
                {
                    DataRow valueRow = dt.NewRow();
                    for (int k = 1; k <= columnCount; k++)
                    {
                        try
                        {
                            valueRow[dt.Columns[k - 1]] = worksheet.Cells[j, k].Value2.ToString();
                        }
                        catch
                        {
                            valueRow[dt.Columns[k - 1]] = "";
                            continue;
                        }

                    }
                    dt.Rows.Add(valueRow);
                }

                //Excel does not need the workbook saved because we are just reading it to transfer to a DataTable
                Excel_CloseExcel(excel, workbook, false);
                return dt;
            }
            else
            {
                return null;
            }
        }
    }
}
