// dotnet add package Z.Dapper.Plus
// dotnet add package Microsoft.Data.SqlClient
// dotnet add package CsvHelper

using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using Z.Dapper.Plus;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;

namespace ExcelToDatabase
{
    public class DynamicModelGenerator
    {
        public static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            string connectionString = "Server=localhost;Database=Store;Trusted_Connection=True;TrustServerCertificate=True;";
            string filePath = "C:/Users/Admin/Desktop/hts_2026_revision_4_xls.xlsx";

            var records = ReadExcel(filePath);

            var modelType = CreateDynamicModel(records[0].Keys);

            var dataList = MapDataToModel(records, modelType);

            string tableName = "DynamicModel";
            CreateTableInDatabase(tableName, modelType, connectionString);

            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                connection.BulkInsert(dataList);
                Console.WriteLine("Дані успішно вставлено в таблицю.");
            }
        }

        static List<Dictionary<string, string>> ReadExcel(string filePath)
        {
            var records = new List<Dictionary<string, string>>();

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.First();
            var headers = worksheet.FirstRowUsed().Cells().Select(c => c.GetValue<string>()).ToList();

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var dict = new Dictionary<string, string>();
                for (int i = 0; i < headers.Count; i++)
                {
                    dict[headers[i]] = row.Cell(i + 1).GetValue<string>();
                }
                records.Add(dict);
            }

            return records;
        }

        static Type CreateDynamicModel(IEnumerable<string> columns)
        {
            var cleanedColumns = columns
                .Select((c, index) => string.IsNullOrWhiteSpace(c) ? $"Column{index + 1}" : c.Trim())
                .ToList();

            var assemblyName = new AssemblyName("DynamicModelAssembly");
            var assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
            var moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");
            var typeBuilder = moduleBuilder.DefineType("DynamicModel", TypeAttributes.Public | TypeAttributes.Class);

            foreach (var column in cleanedColumns)
            {
                var fieldBuilder = typeBuilder.DefineField("_" + column, typeof(string), FieldAttributes.Private);
                var propertyBuilder = typeBuilder.DefineProperty(column, PropertyAttributes.HasDefault, typeof(string), null);

                var getter = typeBuilder.DefineMethod("get_" + column, MethodAttributes.Public, typeof(string), Type.EmptyTypes);
                var getterIl = getter.GetILGenerator();
                getterIl.Emit(OpCodes.Ldarg_0);
                getterIl.Emit(OpCodes.Ldfld, fieldBuilder);
                getterIl.Emit(OpCodes.Ret);

                var setter = typeBuilder.DefineMethod("set_" + column, MethodAttributes.Public, null, new[] { typeof(string) });
                var setterIl = setter.GetILGenerator();
                setterIl.Emit(OpCodes.Ldarg_0);
                setterIl.Emit(OpCodes.Ldarg_1);
                setterIl.Emit(OpCodes.Stfld, fieldBuilder);
                setterIl.Emit(OpCodes.Ret);

                propertyBuilder.SetGetMethod(getter);
                propertyBuilder.SetSetMethod(setter);
            }

            return typeBuilder.CreateType();
        }

        static List<object> MapDataToModel(List<Dictionary<string, string>> records, Type modelType)
        {
            var dataList = new List<object>();
            foreach (var row in records)
            {
                var instance = Activator.CreateInstance(modelType);
                foreach (var kvp in row)
                {
                    var property = modelType.GetProperty(kvp.Key);
                    if (property != null)
                        property.SetValue(instance, kvp.Value);
                }
                dataList.Add(instance);
            }
            return dataList;
        }

        static void CreateTableInDatabase(string tableName, Type modelType, string connectionString)
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();

            var columns = modelType.GetProperties()
                .Select(p => $"[{p.Name}] NVARCHAR(MAX)")
                .ToList();

            string createTableQuery = $@"
                IF OBJECT_ID('{tableName}', 'U') IS NULL
                CREATE TABLE {tableName} (
                    {string.Join(", ", columns)}
                )";

            using var command = new SqlCommand(createTableQuery, connection);
            command.ExecuteNonQuery();
        }
    }
}