using ADOX;
using DataWeb.TurboDB;
using Intuit.MedicalExpenseManager.Db;
using Intuit.MedicalExpenseManager.MemApp;
using System;
using System.IO;
using System.Data;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using System.Collections.Generic;

namespace QuickenMEMExporter {
    internal class Program {
        static void PrintObject(object obj, int level = 0, bool nl = true) {
            Console.Write(new string(' ', level * 4));
            switch (obj) {
                case null: Console.Write("null"); break;
                case bool b: Console.Write(b ? "true" : "false"); break;
                case string s: Console.Write(s); break;
                case int i: Console.Write(i); break;
                case double d: Console.Write(d); break;
                case float f: Console.Write(f); break;
                case long l: Console.Write(l); break;
                case short s: Console.Write(s); break;
                case object[] arr:
                    Console.WriteLine("[");
                    foreach (object o in arr) PrintObject(o, level + 1);
                    Console.Write(new string(' ', level * 4) + "]");
                    break;
                default: Console.Write(obj.ToString()); break;
            }
            if (nl) Console.WriteLine();
        }

        static DataTypeEnum GetType(Type type) {
            if (type == typeof(int)) return DataTypeEnum.adInteger;
            else if (type == typeof(double)) return DataTypeEnum.adDouble;
            else if (type == typeof(string)) return DataTypeEnum.adVarWChar;
            else if (type == typeof(bool)) return DataTypeEnum.adBoolean;
            else if (type == typeof(DateTime)) return DataTypeEnum.adDate;
            else return DataTypeEnum.adEmpty;
        }

        static OleDbType OleType(DataTypeEnum type) {
            switch (type) {
                case DataTypeEnum.adInteger: return OleDbType.Integer;
                case DataTypeEnum.adDouble: return OleDbType.Double;
                case DataTypeEnum.adVarWChar: return OleDbType.VarWChar;
                case DataTypeEnum.adLongVarWChar: return OleDbType.LongVarWChar;
                case DataTypeEnum.adBoolean: return OleDbType.Boolean;
                case DataTypeEnum.adDate: return OleDbType.Date;
                case DataTypeEnum.adDBDate: return OleDbType.DBDate;
                default: return OleDbType.Empty;
            }
        }

        static void Main(string[] args) {
            string path;
            if (args.Length < 1) {
                Console.Write("Enter the path to the database file: ");
                path = Console.ReadLine();
            } else path = args[0];
            FormMain prog = new FormMain();
            switch (prog.DoOpenDataFile(path)) {
                case FormMain.FileOpenStatus.Fail: Console.WriteLine("Failed to open file"); return;
                case FormMain.FileOpenStatus.Cancel: return;
            }
            IDbConnection conn = AppDbConn.GetInstance().Connection;
            Catalog cat = new Catalog();
            Dictionary<string, List<Tuple<string, DataTypeEnum>>> columns = new Dictionary<string, List<Tuple<string, DataTypeEnum>>>();
            try {File.Delete(path + ".accdb");} catch {}
            cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ".accdb" + ";Jet OLEDB:Engine Type=5");
            Console.WriteLine("> Writing table structure...");
            foreach (string name in AppDbConn.GetInstance().Connection.GetTableNames()) {
                string tbl = name.Replace(".dat", "");
                //Console.WriteLine(tbl);
                Table nTable = new Table {
                    Name = tbl
                };
                TurboDBCommand command2 = new TurboDBCommand("SELECT * FROM " + tbl + ";") {
                    Connection = conn
                };
                List<Tuple<string, DataTypeEnum>> cols = new List<Tuple<string, DataTypeEnum>>();
                List<bool> nullable = new List<bool>();
                try {
                    TurboDBDataReader reader = command2.ExecuteReader();
                    DataTable table = reader.GetSchemaTable();
                    foreach (DataRow row in table.Rows) {
                        string colname = row.Field<string>("ColumnName");
                        DataTypeEnum type = GetType(row.Field<Type>("DataType"));
                        int size = row.Field<int>("ColumnSize");
                        if (size > 0xFFFF) size = 0xFFFF;
                        if (type == DataTypeEnum.adVarWChar && size > 255) type = DataTypeEnum.adLongVarWChar;
                        Column col = new Column();
                        col.Name = colname;
                        col.Type = type;
                        col.DefinedSize = size;
                        col.Attributes = (row.Field<bool>("AllowDBNull") && type != DataTypeEnum.adBoolean) ? ColumnAttributesEnum.adColNullable : 0;
                        nTable.Columns.Append(col);
                        cols.Add(new Tuple<string, DataTypeEnum>(colname, type));
                    }
                } catch (Exception e) {
                    Console.WriteLine(e.ToString());
                }
                columns.Add(tbl, cols);
                cat.Tables.Append(nTable);
                Marshal.FinalReleaseComObject(nTable); 
            }
            Marshal.FinalReleaseComObject(cat.Tables);
            Marshal.FinalReleaseComObject(cat.ActiveConnection);
            Marshal.FinalReleaseComObject(cat);
            Console.WriteLine("> Writing table data...");
            using (OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ".accdb")) {
                con.Open();
                foreach (string name in AppDbConn.GetInstance().Connection.GetTableNames()) {
                    string tbl = name.Replace(".dat", "");
                    //Console.WriteLine(tbl);
                    TurboDBCommand command2 = new TurboDBCommand("SELECT * FROM " + tbl + ";") {
                        Connection = conn
                    };
                    try {
                        TurboDBDataReader reader = command2.ExecuteReader();
                        foreach (object obj in reader) {
                            switch (obj) {
                                case object[] arr:
                                    using (OleDbCommand cmd = new OleDbCommand()) {
                                        string com = "INSERT INTO " + tbl + " (";
                                        string param = ") VALUES (";
                                        for (int i = 0; i < columns[tbl].Count; i++) {
                                            if (i > 0) {com += ", "; param += ", ";}
                                            com += "[" + columns[tbl][i].Item1 + "]";
                                            param += "@" + columns[tbl][i].Item1;
                                        }
                                        cmd.Connection = con;
                                        cmd.CommandType = CommandType.Text;
                                        cmd.CommandText = com + param + ")";
                                        Key idx = new Key();
                                        int j = 0;
                                        foreach (object obj2 in arr) {
                                            cmd.Parameters.Add(new OleDbParameter("@" + columns[tbl][j].Item1, OleType(columns[tbl][j].Item2))).Value = obj2 == null ? (columns[tbl][j].Item2 == DataTypeEnum.adBoolean ? false : (object)DBNull.Value) : obj2;
                                            j++;
                                        }
                                        cmd.ExecuteNonQuery();
                                    }
                                    break;
                                default: Console.WriteLine(obj.ToString()); break;
                            }
                        }
                    } catch (Exception e) {
                        Console.WriteLine(e.ToString());
                    }
                }
                Console.WriteLine("> Assigning primary keys...");
                foreach (string name in AppDbConn.GetInstance().Connection.GetTableNames()) {
                    string tbl = name.Replace(".dat", "");
                    foreach (Tuple<string, DataTypeEnum> col in columns[tbl]) {
                        string colname = col.Item1;
                        if (colname == tbl + "Id") {
                            using (OleDbCommand cmd = new OleDbCommand()) {
                                cmd.Connection = con;
                                cmd.CommandType = CommandType.Text;
                                cmd.CommandText = "ALTER TABLE " + tbl + " ADD CONSTRAINT " + tbl + "_" + colname + "_PK PRIMARY KEY (" + colname + ")";
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
                Console.WriteLine("> Creating relationships...");
                foreach (string name in AppDbConn.GetInstance().Connection.GetTableNames()) {
                    string tbl = name.Replace(".dat", "");
                    foreach (Tuple<string, DataTypeEnum> col in columns[tbl]) {
                        string colname = col.Item1;
                        if (colname.EndsWith("Id") && colname != tbl + "Id") {
                            using (OleDbCommand cmd = new OleDbCommand()) {
                                cmd.Connection = con;
                                cmd.CommandType = CommandType.Text;
                                if (colname == "SelectedPlanId") cmd.CommandText = "ALTER TABLE " + tbl + " ADD CONSTRAINT " + tbl + "_SelectedPlanId_FK FOREIGN KEY (SelectedPlanId) REFERENCES InsurancePlan (InsurancePlanId)";
                                else cmd.CommandText = "ALTER TABLE " + tbl + " ADD CONSTRAINT " + tbl + "_" + colname + "_FK FOREIGN KEY (" + colname + ") REFERENCES " + colname.Substring(0, colname.Length - 2) + " (" + colname + ")";
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
            Console.WriteLine("Wrote Access database to " + path + ".accdb");
            AppDbConn.GetInstance().Close();
        }
    }
}
