using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Text;
using static System.Int32;

namespace VFPTableTools
{
    static class Program
    {
        
        static string SelectConnectionString()
        {
            Console.WriteLine("\nSelect a connection option:");
            Console.WriteLine("1. Local (default)");
            Console.WriteLine("2. Test");
            Console.WriteLine("3. Custom");

            Console.Write("\nEnter your choice: ");
            int choice = Parse(Console.ReadLine() ?? string.Empty);

            switch (choice)
            {
                case 1:
                    return @"Provider=VFPOLEDB;Data Source=C:\\Thrive\\sfdata\\vista.dbc;Collating Sequence=machine; Deleted=true;";
                case 2:
                    return @"Provider=VFPOLEDB;Data Source=\\\\DEV-SERVER-1\\Thrive\\sfdata\\vista.dbc;Collating Sequence=machine; Deleted=true;";
                case 3:
                    Console.Write("Enter a custom connection string: ");
                    return Console.ReadLine() ?? string.Empty;
                default:
                    Console.WriteLine("Invalid choice. Using local connection string as default.");
                    return @"Provider=VFPOLEDB;Data Source=C:\\Thrive\\sfdata\\vista.dbc;Collating Sequence=machine; Deleted=true;";
            }
        }
        static void Main()
        {
            var connectionString = SelectConnectionString();
            bool exit = false;

            while (!exit)
            {
                Console.WriteLine("\nMain Menu:");
                Console.WriteLine("1. Create Table");
                Console.WriteLine("2. Table Tools");
                Console.WriteLine("0. Exit");

                Console.Write("\nEnter your choice: ");
                var choice = Parse(Console.ReadLine() ?? string.Empty);

                switch (choice)
                {
                    case 1:
                        CreateTable(connectionString);
                        break;

                    case 2:
                        TableToolsMenu(connectionString);
                        break;

                    case 0:
                        exit = true;
                        break;

                    default:
                        Console.WriteLine("Invalid choice. Please enter a valid option.");
                        break;
                }
            }
        }

        static void CreateTable(string connectionString)
        {
            Console.Write("Enter the table name: ");
            var tableName = Console.ReadLine();

            Console.Write("Enter the number of fields: ");
            var fieldCount = Parse(Console.ReadLine() ?? string.Empty);

            List<string> fields = new List<string>();

            for (int i = 1; i <= fieldCount; i++)
            {
                while (true)
                {
                    Console.Write($"Enter the field name for field #{i}: ");
                    var fieldName = Console.ReadLine();

                    Console.WriteLine("Field Types:");
                    Console.WriteLine("C (Character) - string, stores a fixed-length character string.");
                    Console.WriteLine(
                        "N (Numeric) - decimal, stores a fixed-point number with a specific number of decimal places.");
                    Console.WriteLine("D (Date) - DateTime, stores a date value without time.");
                    Console.WriteLine("T (DateTime) - DateTime, stores a date and time value.");
                    Console.WriteLine("L (Logical) - bool, stores a boolean value (true or false).");
                    Console.WriteLine("Y (Currency) - decimal, stores a currency value with 4 decimal places.");
                    Console.WriteLine("M (Memo) - string, stores variable-length text up to 4 GB.");
                    Console.WriteLine("F (Float) - float, stores approximate numeric values with variable precision.");
                    Console.WriteLine("B (Double) - double, stores approximate numeric values with high precision.");
                    Console.WriteLine("I (Integer) - int, stores 4-byte integer values.");

                    Console.Write($"Enter the field type for field #{i} (C/N/D/T/L/Y/M/F/B/I): ");
                    var fieldType = Console.ReadLine()?.ToUpper();

                    if (fieldType == "C" || fieldType == "N" || fieldType == "F" || fieldType == "B")
                    {
                        Console.Write($"Enter the field length for field #{i}: ");
                        var fieldLength = Parse(Console.ReadLine() ?? string.Empty);
                        fieldType += $"({fieldLength})";
                    }

                    Console.WriteLine($"\nField #{i}: {fieldName} {fieldType}");
                    Console.WriteLine("1. Confirm");
                    Console.WriteLine("2. Edit");
                    Console.Write("Choose an option (1-2): ");
                    int choice = Parse(Console.ReadLine() ?? string.Empty);

                    if (choice == 1)
                    {
                        fields.Add($"{fieldName} {fieldType}");
                        break;
                    }
                }
            }

            Console.WriteLine("\nTable Preview:");
            Console.WriteLine($"Table Name: {tableName}");
            Console.WriteLine("Fields:");
            for (int i = 0; i < fields.Count; i++)
            {
                Console.WriteLine($"Field #{i + 1}: {fields[i]}");
            }

            Console.Write("\nCreate table? (Y/N): ");
            var createTable = Console.ReadLine()?.ToUpper();

            if (createTable == "Y")
            {
#pragma warning disable CA1416
                using OleDbConnection connection = new OleDbConnection(connectionString);

                connection.Open();


                var createTableQuery = $"CREATE TABLE {tableName} ({string.Join(", ", fields)})";
                using (OleDbCommand createTableCommand = new OleDbCommand(createTableQuery, connection))
                {
                    try
                    {
                        createTableCommand.ExecuteNonQuery();
                        Console.WriteLine($"Table '{tableName}' created successfully!");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error creating table: " + ex.Message);
                    }
                }
#pragma warning restore CA1416
            }
        }

        static void TableToolsMenu(string connectionString)
        {
            while (true)
            {
                Console.WriteLine("\nTable Tools Menu");
                Console.WriteLine("================");

                List<string> tableNames = GetTableNames(connectionString);

                if (tableNames.Count == 0)
                {
                    Console.WriteLine("No tables found in database.");
                    return;
                }

                for (int i = 0; i < tableNames.Count; i++)
                {
                    Console.WriteLine($"{i + 1}. {tableNames[i]}");
                }

                Console.Write("\nEnter table name or number (0 to go back): ");

                var input = Console.ReadLine()?.Trim();

                if (TryParse(input, out int tableNumber))
                {
                    if (tableNumber >= 1 && tableNumber <= tableNames.Count)
                    {
                        string tableName = tableNames[tableNumber - 1];
                        TableToolsSubMenu(connectionString, tableName);
                    }
                    else if (tableNumber == 0)
                    {
                        return;
                    }
                    else
                    {
                        Console.WriteLine("Invalid table number.");
                    }
                }
                else
                {
                    if (input != null && tableNames.Contains(input))
                    {
                        TableToolsSubMenu(connectionString, input);
                    }
                    else
                    {
                        Console.WriteLine("Invalid table name.");
                    }
                }
            }
        }

        static List<string> GetTableNames(string connectionString)
        {
            List<string> tableNames = new List<string>();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                DataTable tableSchemaTable =
                    connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] {null, null, null, "TABLE"})!;

                foreach (DataRow row in tableSchemaTable.Rows)
                {
                    string tableName = row["TABLE_NAME"].ToString();
                    tableNames.Add(tableName);
                }
            }

            return tableNames;
        }

        static void TableToolsSubMenu(string connectionString, string tableName)
        {
            while (true)
            {
                Console.WriteLine($"\nTable Tools: {tableName}");
                Console.WriteLine("=====================");
                Console.WriteLine("1. Inspect table");
                Console.WriteLine("2. Export table to CSV");
                Console.WriteLine("3. Mirror table from CSV");
                Console.WriteLine("4. Delete table");
                Console.WriteLine("0. Go back");

                Console.Write("\nEnter option number: ");
                string input = Console.ReadLine().Trim();

                if (TryParse(input, out int option))
                {
                    switch (option)
                    {
                        case 1:
                            InspectTable(connectionString, tableName);
                            break;
                        case 2:
                            Console.Write(
                                "\nEnter the output path (or press Enter to use the root directory -> csv ): ");
                            var path = Console.ReadLine().Trim();
                            ExportTableToCsv(connectionString, tableName, path);
                            break;
                        case 3:
                            Console.Write("\nEnter the name of the primary key column (or press Enter to use 'id'): ");
                            var primaryKeyColumnName = Console.ReadLine().Trim();
                            Console.Write(
                                "\nEnter the path to the CSV file (or press Enter to use the root directory -> csv ): ");
                            var csvPath = Console.ReadLine().Trim();
                            MirrorCsvToTable(connectionString, tableName,
                                string.IsNullOrEmpty(primaryKeyColumnName) ? "id" : primaryKeyColumnName, csvPath);
                            break;
                        case 4:
                            if (DeleteTable(connectionString, tableName))
                            {
                                Console.WriteLine($"The table '{tableName}' was deleted successfully.");
                            }
                            else
                            {
                                Console.WriteLine($"Failed to delete the table '{tableName}'.");
                            }

                            break;
                        case 0:
                            return;
                        default:
                            Console.WriteLine("Invalid option number.");
                            break;
                    }
                }
                else
                {
                    Console.WriteLine("Invalid input.");
                }
            }
        }

        private static bool DeleteTable(string connectionString, string tableName)
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    using (OleDbTransaction transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            using (OleDbCommand command =
                                   new OleDbCommand($"DROP TABLE {tableName}", connection, transaction))
                            {
                                command.ExecuteNonQuery();
                            }

                            transaction.Commit();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"An error occurred while deleting the table: {ex.Message}");
                            transaction.Rollback();
                            return false;
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while connecting to the database: {ex.Message}");
                return false;
            }
        }

        static void InspectTable(string connectionString, string tableName)
        {
            Console.WriteLine($"\nInspecting table '{tableName}':");

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Get the column schema for the selected table
                var inspectTableQuery = $"SELECT * FROM {tableName}";
                using (OleDbCommand inspectTableCommand = new OleDbCommand(inspectTableQuery, connection))
                {
                    using (OleDbDataReader reader = inspectTableCommand.ExecuteReader())
                    {
                        Console.WriteLine($"Table '{tableName}' contains {reader.FieldCount} columns:");
                        Console.WriteLine("Columns:");

                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            string columnName = reader.GetName(i);
                            string columnType = reader.GetDataTypeName(i);
                            int columnLength = reader.GetSchemaTable().Rows[i].Field<int>("ColumnSize");

                            Console.WriteLine($"  {columnName}: {columnType}({columnLength})");
                        }
                    }
                }

                // Get row count
                OleDbCommand rowCountCommand = new OleDbCommand($"SELECT COUNT(*) FROM {tableName}", connection);
                var rowCount = rowCountCommand.ExecuteScalar();
                Console.WriteLine($"Table '{tableName}' contains {rowCount} rows.");
            }
        }

        static void ExportTableToCsv(string connectionString, string tableName, string? outputDirectory = null)
        {
#pragma warning disable CA1416

            Console.Write($"\nExporting table '{tableName}' to CSV... ");

            if (string.IsNullOrEmpty(outputDirectory))
                outputDirectory = "csv";

            Directory.CreateDirectory(outputDirectory);

            // Get the column schema for the selected table
            using OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            var columnSchemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns,
                new object[] {null, null, tableName, null});

            // Build the SELECT query to retrieve all data from the selected table
            StringBuilder selectQueryBuilder = new StringBuilder();
            selectQueryBuilder.Append("SELECT ");

            for (int i = 0; i < columnSchemaTable.Rows.Count; i++)
            {
                string columnName = columnSchemaTable.Rows[i]["COLUMN_NAME"].ToString();
                selectQueryBuilder.Append(columnName);

                if (i < columnSchemaTable.Rows.Count - 1)
                {
                    selectQueryBuilder.Append(", ");
                }
            }

            selectQueryBuilder.Append($" FROM {tableName} ORDER BY 1");

            // Execute the SELECT query to retrieve all data from the selected table
            OleDbCommand selectCommand = new OleDbCommand(selectQueryBuilder.ToString(), connection);
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(selectCommand);

            DataTable dataTable = new DataTable();
            dataAdapter.Fill(dataTable);

            // Write the data to a CSV file
            using (StreamWriter writer = new StreamWriter(Path.Combine(outputDirectory, $"{tableName}.csv")))
            {
                // Write the header row
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    writer.Write(dataTable.Columns[i].ColumnName);

                    if (i < dataTable.Columns.Count - 1)
                    {
                        writer.Write(",");
                    }
                }

                writer.WriteLine();

                // Write the data rows
                foreach (DataRow row in dataTable.Rows)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        string value = row[i].ToString();
                        if (value.Contains(","))
                        {
                            value = $"\"{value}\"";
                        }

                        writer.Write(value);

                        if (i < dataTable.Columns.Count - 1)
                        {
                            writer.Write(",");
                        }
                    }

                    writer.WriteLine();
                }
            }

            //get current directory

            Console.WriteLine($"Exported to CSV -> {outputDirectory}\\{tableName}.csv");
        }

        static void MirrorCsvToTable(string connectionString, string tableName, string primaryKeyColumn,
            string? csvFilePath = null)
        {
            if (string.IsNullOrEmpty(csvFilePath))
                csvFilePath = $"csv\\{tableName}.csv";

            if (!File.Exists(csvFilePath))
            {
                Console.WriteLine($"CSV file does not exist: {csvFilePath}");
                return;
            }

            try
            {
                DataTable schemaTable;
                // Fetch the schema from the Visual FoxPro table
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    schemaTable = connection.GetSchema("Columns", new string[] {null, null, tableName, null});
                    connection.Close();
                }

                if (schemaTable.Rows.Count == 0)
                {
                    Console.WriteLine($"Error: Table {tableName} not found");
                    return;
                }

                HashSet<object> primaryKeysInTable;

                // Read the primary key values from the Visual FoxPro table
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command =
                           new OleDbCommand($"SELECT {primaryKeyColumn} FROM {tableName}", connection))
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        primaryKeysInTable = new HashSet<object>();
                        while (reader.Read())
                        {
                            primaryKeysInTable.Add(reader.GetValue(0));
                        }
                    }

                    connection.Close();
                }

                // Read the CSV file into a DataTable
                DataTable csvData = new DataTable();

                // Create columns based on schema
                foreach (DataRow schemaRow in schemaTable.Rows)
                {
                    string columnName = schemaRow["COLUMN_NAME"].ToString();
                    OleDbType oleDbType = (OleDbType) schemaRow["DATA_TYPE"];
                    Type columnType = OleDbTypeToType(oleDbType);
                    csvData.Columns.Add(columnName, columnType);
                }

                using (StreamReader sr = new StreamReader(csvFilePath))
                {
                    sr.ReadLine(); // Skip the header row

                    while (!sr.EndOfStream)
                    {
                        var line = sr.ReadLine();
                        var rows = new List<string>();
                        var inQuotes = false;
                        var currentValue = new StringBuilder();

                        foreach (char c in line)
                        {
                            if (c == '\"')
                            {
                                inQuotes = !inQuotes;
                            }
                            else if (c == ',' && !inQuotes)
                            {
                                rows.Add(currentValue.ToString());
                                currentValue.Clear();
                            }
                            else
                            {
                                currentValue.Append(c);
                            }
                        }

                        rows.Add(currentValue.ToString());
                        DataRow newRow = csvData.NewRow();

                        bool isEmptyRow = true;
                        for (int i = 0; i < schemaTable.Rows.Count; i++)
                        {
                            Type columnType = newRow.Table.Columns[i].DataType;
                            string trimmedValue = rows[i].Trim();

                            if (!string.IsNullOrEmpty(trimmedValue))
                            {
                                isEmptyRow = false;
                                try
                                {
                                    newRow[i] = Convert.ChangeType(trimmedValue, columnType);
                                }
                                catch (FormatException)
                                {
                                    Console.WriteLine(
                                        $"Error: The input string '{rows[i]}' at column {i + 1} was not in a correct format.");
                                    // Handle the error, for example, by skipping the row or setting a default value.
                                    newRow[i] = GetDefaultValueForType(columnType);
                                }
                            }
                            else
                            {
                                newRow[i] = GetDefaultValueForType(columnType);
                            }
                        }

                        if (isEmptyRow)
                        {
                            continue;
                        }

                        // Process the row
                        object primaryKeyValue = newRow[primaryKeyColumn];
                        if (primaryKeysInTable.Contains(primaryKeyValue))
                        {
                            // Update the row in the Visual FoxPro table
                            UpdateRow(connectionString, tableName, newRow, primaryKeyColumn);
                            primaryKeysInTable.Remove(primaryKeyValue);
                        }
                        else
                        {
                            // Insert the row into the Visual FoxPro table
                            InsertRow(connectionString, tableName, newRow);
                        }
                    }
                }

                // Delete remaining rows from the Visual FoxPro table
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand())
                    {
                        command.Connection = connection;

                        foreach (object primaryKeyValue in primaryKeysInTable)
                        {
                            command.CommandText = $"DELETE FROM {tableName} WHERE {primaryKeyColumn} = ?";
                            command.Parameters.Clear();
                            command.Parameters.AddWithValue($"@{primaryKeyColumn}", primaryKeyValue);
                            command.ExecuteNonQuery();
                            Console.WriteLine($"Deleted row with {primaryKeyColumn} = {primaryKeyValue}");
                        }
                    }

                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static void UpdateRow(string connectionString, string tableName, DataRow row, string primaryKeyColumn)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbCommand command = new OleDbCommand())
                {
                    command.Connection = connection;

                    var setClause = new StringBuilder();
                    foreach (DataColumn column in row.Table.Columns)
                    {
                        if (column.ColumnName != primaryKeyColumn)
                        {
                            setClause.Append(column.ColumnName).Append(" = ?").Append(",");
                        }
                    }

                    setClause.Length--; // Remove the trailing comma

                    command.CommandText = $"UPDATE {tableName} SET {setClause} WHERE {primaryKeyColumn} = ?";
                    command.Parameters.Clear();
                    foreach (DataColumn column in row.Table.Columns)
                    {
                        if (column.ColumnName != primaryKeyColumn)
                        {
                            command.Parameters.AddWithValue($"@{column.ColumnName}", row[column]);
                        }
                    }

                    command.Parameters.AddWithValue($"@{primaryKeyColumn}", row[primaryKeyColumn]);
                    command.ExecuteNonQuery();
                }

                connection.Close();
            }

            Console.WriteLine($"Updated row for primary key value {row[primaryKeyColumn]}");
        }

        private static void InsertRow(string connectionString, string tableName, DataRow row)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbCommand command = new OleDbCommand())
                {
                    command.Connection = connection;

                    var columns = new StringBuilder();
                    var values = new StringBuilder();

                    foreach (DataColumn column in row.Table.Columns)
                    {
                        columns.Append(column.ColumnName).Append(",");
                        values.Append("?").Append(",");
                    }

                    columns.Length--; // Remove the trailing comma
                    values.Length--; // Remove the trailing comma

                    string insertCommand = $"INSERT INTO {tableName} ({columns}) VALUES ({values})";
                    command.CommandText = insertCommand;

                    command.Parameters.Clear();
                    foreach (DataColumn column in row.Table.Columns)
                    {
                        command.Parameters.AddWithValue($"@{column.ColumnName}", row[column]);
                    }

                    command.ExecuteNonQuery();
                }

                connection.Close();
            }

            Console.WriteLine($"Inserted row for primary key value {row[0]}");
        }


        private static string ConvertToVfpValueString(object value, string vfpDataType)
        {
            switch (vfpDataType)
            {
                case "10": // Character
                    return "'" + ((string) value).Replace("'", "''") + "'";
                case "3": // Numeric
                    return ((decimal) value).ToString(CultureInfo.InvariantCulture);
                case "11": // Date
                    return "CTOD('" + ((DateTime) value).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture) + "')";
                case "12": // DateTime
                    return "DATETIME('" +
                           ((DateTime) value).ToString("MM/dd/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture) + "')";
                case "9": // Currency
                    return ((decimal) value).ToString("C2", CultureInfo.InvariantCulture);
                case "2": // Integer
                case "4": // Single
                case "5": // Double
                case "6": // Float
                    return ((IConvertible) value).ToString(CultureInfo.InvariantCulture);
                case "7": // DateTimeStamp
                    return "DATETIMESTAMP('" +
                           ((DateTime) value).ToString("MM/dd/yyyy hh:mm:ss.fff tt", CultureInfo.InvariantCulture) +
                           "')";
                default:
                    return "'" + value.ToString().Replace("'", "''") + "'";
            }
        }


        private static object? GetDefaultValueForType(Type columnType)
        {
            if (columnType == typeof(string))
                return string.Empty;

            return columnType.IsValueType ? Activator.CreateInstance(columnType)! : null;
        }

        private static Type OleDbTypeToType(OleDbType oleDbType)
        {
            switch (oleDbType)
            {
                case OleDbType.Numeric:
                    return typeof(decimal);
                case OleDbType.Boolean:
                    return typeof(bool);
                case OleDbType.TinyInt:
                    return typeof(sbyte);
                case OleDbType.UnsignedTinyInt:
                    return typeof(byte);
                case OleDbType.SmallInt:
                    return typeof(short);
                case OleDbType.UnsignedSmallInt:
                    return typeof(ushort);
                case OleDbType.Integer:
                    return typeof(int);
                case OleDbType.UnsignedInt:
                    return typeof(uint);
                case OleDbType.BigInt:
                    return typeof(long);
                case OleDbType.UnsignedBigInt:
                    return typeof(ulong);
                case OleDbType.Single:
                    return typeof(float);
                case OleDbType.Double:
                    return typeof(double);
                case OleDbType.Currency:
                    return typeof(decimal);
                case OleDbType.Date:
                case OleDbType.DBTimeStamp:
                    return typeof(DateTime);
                case OleDbType.BSTR:
                case OleDbType.Char:
                case OleDbType.VarChar:
                case OleDbType.LongVarChar:
                case OleDbType.WChar:
                case OleDbType.VarWChar:
                case OleDbType.LongVarWChar:
                    return typeof(string);
                default:
                    throw new ArgumentOutOfRangeException(nameof(oleDbType), oleDbType, "Unsupported OleDbType");
            }
        }
    }

}