/*
Json Excel Converter

by ray@raymix.net


MIT License

Copyright (c) 2018 RayMix.net

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JsonExcel
{
    class Program
    {

        static bool toJson = false;
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                ShowHelp();
                return;
            }

            bool force = false;

            string type = args[0];
            if (args[0] != "json" && args[0] != "excel")
            {
                ShowHelp();
                return;
            }
            toJson = type == "json";
            string source = args[1];
            string dest = "";
            if (args[2] == "force")
                force = true;
            else
                dest = args[2];
            if (args.Length > 3 && args[3] == "force") force = true;

            if (File.Exists(source))
                ConvertFile(source, dest, force);
            else if (Directory.Exists(source))
                ConvertDirectory(source, dest, force);
            else
                ShowHelp();
        }

        static void ConvertDirectory(string source,string destination,bool force)
        {
            string[] sourceFiles = Directory.GetFiles(source,"*", SearchOption.AllDirectories).Where(f => toJson ? f.EndsWith(".xlsx") : (f.EndsWith(".txt") || f.EndsWith(".txt"))).ToArray(); ;
            foreach (string file in sourceFiles)
            {
                string target = file;
                if (destination != "")target = file.Replace(source, destination);
                ConvertFile(file, target, force);
            }
        }

        static void ConvertFile(string source, string destination, bool force)
        {
            if(destination == "")
                destination = source;

            if (toJson)
                destination = destination.Replace(".xlsx", ".txt");
            else
            {
                destination = destination.Replace(".txt", ".xlsx");
                destination = destination.Replace(".json", ".xlsx");
            }

            if (!toJson && !force && File.Exists(destination))return;

            try
            {
                if (toJson)
                    XlsxToJson(source, destination);
                else
                    JsonToXlsx(source, destination);

                Console.WriteLine(destination + " [OK]");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(destination + " [FAILED]");
                Console.WriteLine(ex.Message);
                Console.ResetColor();

            }
        }

        static int type_row = 2;
        static int header_row = 3;
        static int row = 4;
        static int col = 1;
        static int max_col = 1;
        static int key_index = 0;

        static void JsonToXlsx(string json_file, string xlsx_file)
        {
            row = header_row + 1;
            col = 1;
            max_col = 1;
            string json = File.ReadAllText(json_file);
            object obj = JsonConvert.DeserializeObject(json);

            if (File.Exists(xlsx_file)) File.Delete(xlsx_file);
            using (var excel = new ExcelPackage(new FileInfo(xlsx_file)))
            {
                var ws = excel.Workbook.Worksheets.Add(Path.GetFileNameWithoutExtension(xlsx_file));

                if (obj is JObject)
                {
                    JObject json_obj = (JObject)obj;

                    ws.Cells[header_row, col++].Value = "Key";
                    foreach (JToken token in json_obj.Children())
                    {
                        key_index = 0;
                        col = 1;
                        JProperty prop = (JProperty)token;
                        ws.Cells[row, col++].Value = prop.Name;
                        ParseJsonToken("", prop.Value, ws);
                        row++;
                        max_col = Math.Max(col, max_col);
                    }
                }
                else if(obj is JArray)
                {
                    JArray json_obj = (JArray)obj;
                    ws.Cells[header_row, col++].Value = "Array";

                    foreach (JToken token in json_obj.Children())
                    {
                        key_index = 0;
                        col = 1;
                        ParseJsonToken("", token, ws);
                        row++;
                        max_col = Math.Max(col, max_col);
                    }

                }
                else
                {

                }

                for (int i = 1; i <= max_col; i++)
                {
                    if (ws.Column(i) != null)
                        ws.Column(i).AutoFit();
                }
                ws.Tables.Add(new ExcelAddress(header_row, 1, row - 1, max_col-1), "Medium6");
                excel.Save();
            }
        }

        static void XlsxToJson(string xlsx_file, string json_file)
        {
            using (var excel = new ExcelPackage(new FileInfo(xlsx_file)))
            {
                var ws = excel.Workbook.Worksheets[1];

                if (ws.Cells[header_row, 1].Value.Equals("Key") && ws.Tables[0].Address.Columns > 2)
                {
                    JObject obj = new JObject();
                    Dictionary<int, JObject> Owners = new Dictionary<int, JObject>();
                    for (int r = ws.Tables[0].Address.Start.Row + 1; r <= ws.Tables[0].Address.Rows + ws.Tables[0].Address.Start.Row - 1; r++)
                    {
                        if (ws.Cells[r, 1].Value != null)
                        {
                            JObject item = new JObject();
                            obj.Add(ws.Cells[r, 1].Value.ToString(), (JToken)item);
                            Owners[1] = item;
                        }

                        JObject Owner = null;
                        for (int c = ws.Tables[0].Address.Start.Column + 1; c <= ws.Tables[0].Address.Columns; c++)
                        {
                            if (Owners.ContainsKey(c - 1))
                                Owner = Owners[c - 1];
                            string header = ws.Cells[header_row, c].Value.ToString();
                            if (ws.Cells[type_row, c].Value == null || header == "Value")
                            {
                                if (ws.Cells[r, c].Value != null)
                                {
                                    try
                                    {
                                        if (ws.Cells[header_row, c + 1].Value != null && ws.Cells[header_row, c + 1].Value.ToString() == "Value")
                                        {//KVOBJECT
                                            JTokenType vtype = GetJTokenType(ws.Cells[r, c + 1].Value.ToString());
                                            JToken v = GetJToken(vtype, ws.Cells[r, c + 1].Value);
                                            if (v != null) Owner.Add(ws.Cells[r, c].Value.ToString(), v);
                                            c++;
                                        }
                                        else
                                        {
                                            JObject subItem = new JObject();
                                            Owner.Add(ws.Cells[r, c].Value.ToString(), subItem);
                                            Owners[c] = subItem;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception(string.Format("Row:{0},Col:{1} ERROR:{2}", r, c, ex.Message), ex.InnerException);
                                    }
                                }
                                continue;
                            }

                            if (ws.Cells[type_row, c].Value.ToString() != "EditOnly")
                            {
                                try
                                {
                                    JTokenType type = (JTokenType)Enum.Parse(typeof(JTokenType), ws.Cells[type_row, c].Value.ToString());
                                    JToken val = GetJToken(type, ws.Cells[r, c].Value);
                                    if (val != null) Owner.Add(header, val);
                                    if (header.Contains("."))
                                    {
                                        //JArray array = new JArray();
                                    }
                                }
                                catch(Exception ex)
                                {
                                    throw new Exception(string.Format("Row:{0},Col:{1} ERROR:{2}", r, c, ex.Message), ex.InnerException);
                                }
                            }
                        }
                    }
                    File.WriteAllText(json_file, JsonConvert.SerializeObject(obj, Formatting.Indented));
                }
                else if (ws.Cells[header_row, 1].Value.Equals("Value")) 
                {//ARRAY
                    JArray array = new JArray();
                    for (int r = ws.Tables[0].Address.Start.Row + 1; r <= ws.Tables[0].Address.Rows + ws.Tables[0].Address.Start.Row - 1; r++)
                    {
                        if (ws.Cells[r, 1].Value != null)
                        {
                            JTokenType type = (JTokenType)Enum.Parse(typeof(JTokenType), ws.Cells[type_row, 1].Value.ToString());
                            JToken val = GetJToken(type, ws.Cells[r, 1].Value);
                            if (val != null) array.Add(GetJToken(type, val));
                        }
                    }
                    File.WriteAllText(json_file, JsonConvert.SerializeObject(array, Formatting.Indented));
                }
                else if (ws.Cells[header_row, 2].Value.Equals("Value"))
                {//JObject (Key: Value)
                    JObject obj = new JObject();
                    for (int r = ws.Tables[0].Address.Start.Row + 1; r <= ws.Tables[0].Address.Rows + ws.Tables[0].Address.Start.Row - 1; r++)
                    {
                        if (ws.Cells[r, 2].Value != null)
                        {
                            JTokenType type = GetJTokenType(ws.Cells[r, 2].Value.ToString());
                            JToken val = GetJToken(type, ws.Cells[r, 2].Value);
                            if (val != null) obj.Add(ws.Cells[r, 1].Value.ToString(), val);
                        }
                    }
                    File.WriteAllText(json_file, JsonConvert.SerializeObject(obj, Formatting.Indented));

                }
            }
        }

        static JTokenType GetJTokenType(string value)
        {
            JTokenType type = IsNumberic(value) ? IsFloat(value) ? JTokenType.Float : JTokenType.Integer : JTokenType.String;
            if (type == JTokenType.String)
            {
                if (value.StartsWith("["))
                    type = JTokenType.Array;
                if (value.Equals("false", StringComparison.CurrentCultureIgnoreCase) || value.Equals("true", StringComparison.CurrentCultureIgnoreCase))
                    type = JTokenType.Boolean;
            }
            return type;
        }

        static JToken GetJToken(JTokenType type, object value)
        {
            if (value == null)
                return null;
            switch (type)
            {
                case JTokenType.Array:
                    return (JToken)JsonConvert.DeserializeObject(value.ToString());
                case JTokenType.Float:
                    return new JValue(float.Parse(value.ToString()));
                case JTokenType.Integer:
                    return new JValue(int.Parse(value.ToString()));
                case JTokenType.Boolean:
                    return new JValue(bool.Parse(value.ToString()));
                default:
                    return new JValue(value.ToString());
            }
        }

        static void ParseJsonObject(string parentKey, JObject obj,ExcelWorksheet ws)
        {
            foreach (JToken token in obj.Children())
            {
                ParseJsonToken(parentKey,token,ws);
            }
            //Console.WriteLine(obj.ToString());
        }

        static int FindColumnNo(ExcelWorksheet ws,string name)
        {
            int i = 1;
            while(true)
            {
                object val = ws.Cells[header_row, i].Value;
                if (val==null || val.ToString() =="")
                    return 0;
                if (val.Equals(name))
                    return i;
                i++;
            }
        }

        static void SetType(ExcelWorksheet ws, JTokenType type, int column)
        {
            if (ws.Cells[type_row, column].Value == null)
            {
                ws.Cells[type_row, column].Value = type;
                return;
            }

            if (ws.Cells[type_row, column].Value.Equals(JTokenType.String))
                return;

            if (ws.Cells[type_row, column].Value.Equals(JTokenType.Float))
                return;

            ws.Cells[type_row, column].Value = type;
        }
        static void ParseJsonToken(string parentKey, JToken token, ExcelWorksheet ws)
        {
            switch (token.Type)
            {
                case JTokenType.Property:
                    {
                        JProperty prop = (JProperty)token;
                        string name = parentKey == "" ? prop.Name : parentKey + "." + prop.Name;
                        if (ws.Cells[header_row, col].Value == null || !ws.Cells[header_row, col].Value.Equals(name))
                        {//列不一致
                            int findCol = FindColumnNo(ws, name);
                            if (findCol > 0)
                            {
                                col = findCol;
                            }
                            else
                            {
                                if (ws.Cells[header_row, col].Value != null)
                                    ws.InsertColumn(col, 1);
                                ws.Cells[header_row, col].Value = name;
                            }
                        }

                        if (prop.Value.Type == JTokenType.Object)
                        {
                            ParseJsonObject(name, (JObject)prop.Value, ws);
                        }
                        else if (prop.Value.Type == JTokenType.Array)
                        {
                            SetType(ws, prop.Value.Type, col);
                            ws.Cells[type_row, col].Value = prop.Value.Type;
                            ws.Cells[row, col++].Value = prop.Value;
                        }
                        else
                        {
                            SetType(ws, prop.Value.Type, col);
                            ws.Cells[row, col++].Value = prop.Value;
                        }
                    }
                    break;
                case JTokenType.Object:
                    {
                        JObject obj = (JObject)token;

                        int orign_col = col;
                        int orign_row = row;
                        foreach (JToken tk in obj.Children())
                        {
                            JProperty prop = (JProperty)tk;
                            int o;
                            int.TryParse(prop.Name, out o);
                            if(int.TryParse(prop.Name, out o) && prop.Name == o.ToString() || prop.Value.Type == JTokenType.Object)
                            {
                                col = orign_col;
                                key_index++;
                                if (ws.Cells[header_row, col].Value == null) ws.Cells[header_row, col].Value = "Key" + key_index.ToString();
                                ws.Cells[row, col++].Value = prop.Name;
                                ParseJsonToken("", prop.Value, ws);
                                row++;
                                max_col = Math.Max(col, max_col);
                            }
                            else
                            {
                                ParseJsonToken(parentKey, tk, ws);
                            }
                        }
                        if (row > orign_row)
                            row--;
                    }
                    break;
                default:
                    {
                        JValue val = (JValue)token;
                        SetType(ws, val.Type, col);
                        ws.Cells[header_row, col].Value = "Value";
                        ws.Cells[row, col++].Value = val.Value;
                    }
                    break;
            }
            max_col = Math.Max(col, max_col);
        }



        static bool IsNumberic(string text)
        {
            foreach (char c in text)
            {
                if (!char.IsDigit(c))
                {
                    if (c != '.' && c != '-')
                        return false;
                }
            }
            return true;
        }

        static bool IsFloat(string text)
        {
            return (IsNumberic(text) && text.IndexOf(".") > 0);
        }

        static void ShowHelp()
        {
            Console.WriteLine(@"Json Excel Converter

by ray@raymix.net

json-excel <json|excel> source [destination] [force]
    json        : convert Excel format to Json.
    excel       : convert Json format to Excel.
    source      : Specify input file or directory to convert.
    destination : Specify output file or directory,
                  same as source if unspecified
    force       : force overwrite excel file

https://github.com/rayturn/json-excel
");
        }
    }
}
