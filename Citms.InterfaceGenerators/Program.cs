using Newtonsoft.Json;
using Novacode;
using Swashbuckle.Swagger;
using System;
using System.Timers;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Net;

namespace Citms.InterfaceGenerators
{
    class Program
    {
        private static Dictionary<string, string> ModuleDict = new Dictionary<string, string>()
        {
            { "Capture" ,"设备抓拍" },
            { "Common" ,"公共接口" },
            { "Intercepting" ,"布控管理" },
            { "Maintenance" ,"运维管理" },
            { "MessageCenter" ,"消息中心" },
            { "Passport" ,"通行证" },
            { "Pis" ,"地图模块" },
            { "Police" ,"勤务管理" },
            { "Punish" ,"违法模块" },
            {"Rules" ,"违法模块" },
            {"SchemeRestriction" ,"违法模块" },
            { "Reidentification" ,"二次识别" },
            { "Report" ,"数据统计" },
            { "Security" ,"系统管理" },
            { "SysManage" ,"基础数据" },
            { "Track","轨迹分析"},
            {"TrafficMonitor","通行记录" },
            {"VehicleInfo","违法模块" },
            {"VideoHistory","通行记录" }
        };


        static void Main(string[] args)
        {
            //http://localhost:1371/swagger/docs/v1

            Console.WriteLine("请输入swagger接口地址(http://192.168.0.133:6001/swagger/v1/swagger.json)：");
            string url = Console.ReadLine().ToLower();
            while (true)
            {
                if (!url.StartsWith("http://") && !url.StartsWith("https://"))
                {
                    Console.WriteLine("请输入正确的接口地址(http://192.168.0.133:6001/swagger/v1/swagger.json)：");
                    url = Console.ReadLine().ToLower();
                }
                else
                {
                    break;
                }
            }
            Console.WriteLine("开始从IMS中获取接口描述");
            string desc = GetInterfaceDesc(url).Replace("$ref", "refDef");
            Console.WriteLine("获取接口描述已完成，开始生成接口文档");
            CreateApiDoc(desc);

            Console.WriteLine("接口文档已经完成到【" + AppDomain.CurrentDomain.BaseDirectory + "API说明手册.docx" + "】路径下");
            Console.Read();
        }

        /// <summary>
        /// 获取IMS接口描述
        /// </summary>
        /// <param name="url">IMS站点地址</param>
        /// <returns>SwaggerDocument描述</returns>
        public static string GetInterfaceDesc(string url)
        {
            HttpWebRequest req = null;
            HttpWebResponse res = null;
            try
            {
                req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "GET";
                req.Timeout = 200000;
                res = (HttpWebResponse)req.GetResponse();
                using (StreamReader sr = new StreamReader(res.GetResponseStream(), Encoding.UTF8))
                {
                    return sr.ReadToEnd();
                }
            }
            catch (WebException exception)
            {
                Console.WriteLine("获取IMS接口描述异常:" + exception.ToString());
                throw exception;
            }
            finally
            {
                if (res != null)
                {
                    res.Close();
                    res = null;
                }
                if (req != null)
                {
                    req.Abort();
                    req = null;
                }
            }
        }

        public static void CreateApiDoc(string apiDesc)
        {
            int minRowHeight = 28;
            int fontSize = 11;
            SwaggerDocument data = JsonConvert.DeserializeObject<SwaggerDocument>(apiDesc);
            if (data != null)
            {
                Dictionary<string, dynamic> dict = new Dictionary<string, dynamic>();
                foreach (var key in data.definitions.Keys)
                {
                    InitParamsDefinedCache(data.definitions[key].properties, "#/definitions/" + key, dict, data.definitions);
                }

                //生成word文档
                using (DocX doc = DocX.Create(AppDomain.CurrentDomain.BaseDirectory + "API说明手册.docx", DocumentTypes.Document))
                {
                    Paragraph pSumHead = doc.InsertParagraph();
                    pSumHead.Alignment = Alignment.center;
                    pSumHead.Append("第四章接口设计规范").Bold().Color(Color.Black).FontSize(fontSize + 8).Heading(HeadingType.Heading1);
                    pSumHead.LineSpacing = 1.5f;

                    var dictController = data.ControllerDesc==null?new Dictionary<string,string>(): data.ControllerDesc;

                    string curHeader = string.Empty, preHeader = string.Empty;
                    int oneIndex = 1;
                    int everyCount = data.paths.Keys.Count / 10, curIndex = 1;
                    foreach (string path in data.paths.Keys)
                    {
                        var pathItem = data.paths[path];
                        if (curIndex % everyCount == 0)
                        {
                            Console.WriteLine("接口文档生成进度" + curIndex * 10 / everyCount + "%");
                        }
                        curIndex++;
                        string method = string.Empty;
                        var Operation = GetOper(pathItem, out method);
                        string controllerName = Operation.tags[0];
                        Paragraph pHead = doc.InsertParagraph();
                        if (dictController.ContainsKey(controllerName))
                        {
                            curHeader = dictController[controllerName];
                        }
                        else
                        {
                            curHeader = controllerName;
                        }
                        if (curHeader != preHeader)
                        {
                            pHead.LineSpacing = 1.5f;
                            pHead.Append("4." + oneIndex + " " + curHeader).Color(Color.Black).Bold().FontSize(fontSize + 4).Heading(HeadingType.Heading2);
                            preHeader = curHeader;
                            oneIndex++;
                        }

                        Table table = doc.InsertTable(20, 3);
                        table.Design = TableDesign.TableGrid;
                        table.Alignment = Alignment.center;
                        List<Row> rows = table.Rows;
                        //首行接口描述
                        Row rowHeader = rows[0];
                        rowHeader.Height = minRowHeight;
                        rowHeader.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                        rowHeader.Cells[0].FillColor = Color.FromArgb(214, 227, 188);
                        Paragraph p = rowHeader.Cells[0].Paragraphs[0];
                        p.Alignment = Alignment.left;
                        p.Append("接口描述").Bold().FontSize(fontSize);
                        rowHeader.MergeCells(1, 2);
                        rowHeader.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                        rowHeader.Cells[1].FillColor = Color.FromArgb(214, 227, 188);
                        p = rowHeader.Cells[1].Paragraphs[0];
                        p.Alignment = Alignment.left;

                        p.Append(Operation.summary).FontSize(fontSize);
                        rowHeader.Cells[0].Width = 110;
                        rowHeader.Cells[1].Width = 420;

                        //第二行接口地址
                        Row newRow = rows[1];
                        newRow.Height = minRowHeight;
                        newRow.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                        p = newRow.Cells[0].Paragraphs[0];
                        p.Alignment = Alignment.left;
                        p.Append("接口地址").FontSize(fontSize);
                        newRow.MergeCells(1, 2);
                        newRow.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                        p = newRow.Cells[1].Paragraphs[0];
                        p.Append(path).FontSize(fontSize);

                        //第三行请求方法
                        newRow = rows[2];
                        newRow.Height = minRowHeight;
                        newRow.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                        p = newRow.Cells[0].Paragraphs[0];
                        p.Alignment = Alignment.left;
                        p.Append("请求方法").FontSize(fontSize);
                        newRow.MergeCells(1, 2);
                        newRow.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                        p = newRow.Cells[1].Paragraphs[0];
                        p.Append(method).FontSize(fontSize);

                        //第四行请求参数
                        newRow = rows[3];
                        newRow.Height = minRowHeight;
                        newRow.MergeCells(0, 2);
                        newRow.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                        newRow.Cells[0].FillColor = Color.FromArgb(214, 227, 188);
                        p = newRow.Cells[0].Paragraphs[0];
                        p.Alignment = Alignment.left;
                        p.Append("请求参数").Bold().FontSize(fontSize);

                        //第五行请求参数表头
                        newRow = rows[4];
                        newRow.Height = minRowHeight;
                        newRow.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                        newRow.Cells[0].FillColor = Color.FromArgb(217, 217, 217);
                        p = newRow.Cells[0].Paragraphs[0];
                        p.Alignment = Alignment.center;
                        p.Append("参数名").FontSize(fontSize);
                        newRow.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                        newRow.Cells[1].FillColor = Color.FromArgb(217, 217, 217);
                        p = newRow.Cells[1].Paragraphs[0];
                        p.Alignment = Alignment.center;
                        p.Append("描述").FontSize(fontSize);
                        newRow.Cells[2].VerticalAlignment = VerticalAlignment.Center;
                        newRow.Cells[2].FillColor = Color.FromArgb(217, 217, 217);
                        p = newRow.Cells[2].Paragraphs[0];
                        p.Alignment = Alignment.center;
                        p.Append("示例").FontSize(fontSize);

                        //第六行开始循环写入参数
                        int index = 5;
                        if (Operation.parameters != null)
                        {
                            foreach (var item in Operation.parameters)
                            {
                                if (item.@in == "path")
                                {
                                    newRow = rows[index];
                                    newRow.Height = minRowHeight;
                                    newRow.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                    p = newRow.Cells[0].Paragraphs[0];
                                    p.Append(item.name).FontSize(fontSize);

                                    newRow.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                                    p = newRow.Cells[1].Paragraphs[0];
                                    p.Append(item.description).FontSize(fontSize);

                                    newRow.Cells[2].VerticalAlignment = VerticalAlignment.Center;
                                    p = newRow.Cells[2].Paragraphs[0];
                                    p.Append(item.type).FontSize(fontSize);

                                    index++;
                                }
                                else if (item.@in == "body")
                                {
                                    newRow = rows[index];
                                    newRow.Height = minRowHeight;
                                    newRow.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                    p = newRow.Cells[0].Paragraphs[0];
                                    p.Append(item.name).FontSize(fontSize);


                                    newRow.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                                    p = newRow.Cells[1].Paragraphs[0];
                                    string refDef = item.schema.refDef;
                                    dynamic pp;
                                    if (!string.IsNullOrEmpty(refDef) && dict.TryGetValue(refDef, out pp))
                                    {
                                        p.Append(JsonConvert.SerializeObject(pp, Newtonsoft.Json.Formatting.Indented)).FontSize(fontSize);
                                        newRow.Height = 10 * minRowHeight;
                                    }

                                    newRow.Cells[2].VerticalAlignment = VerticalAlignment.Center;
                                    p = newRow.Cells[2].Paragraphs[0];
                                    p.Append(item.description).FontSize(fontSize);

                                    index++;
                                }
                            }
                        }

                        //写入返回值
                        newRow = rows[index];
                        newRow.Height = minRowHeight;
                        newRow.MergeCells(0, 2);
                        newRow.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                        newRow.Cells[0].FillColor = Color.FromArgb(214, 227, 188);
                        p = newRow.Cells[0].Paragraphs[0];
                        p.Alignment = Alignment.left;
                        p.Append("返回值").Bold().FontSize(fontSize);
                        index++;

                        newRow = rows[index];
                        newRow.Height = 10 * minRowHeight;
                        newRow.MergeCells(0, 2);
                        newRow.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                        p = newRow.Cells[0].Paragraphs[0];
                        p.Alignment = Alignment.left;

                        if (Operation.responses.ContainsKey("200"))
                        {
                            string refDefKey = Operation.responses["200"].schema?.refDef;
                            dynamic pp1;
                            if (!string.IsNullOrEmpty(refDefKey) && dict.TryGetValue(refDefKey, out pp1))
                            {
                                p.Append(JsonConvert.SerializeObject(pp1, Newtonsoft.Json.Formatting.Indented)).FontSize(fontSize);
                                newRow.Height = 10 * minRowHeight;
                            }
                        }
                        else
                        {
                            p.Append("无").FontSize(fontSize);
                        }
                        index++;

                        int i = index;
                        while (index < 20)
                        {
                            table.RemoveRow(i);
                            index++;
                        }
                    }

                    doc.Save();
                }
            }
        }

        private static Operation GetOper(PathItem item, out string method)
        {
            Operation opr = null;
            if (item.get != null)
            {
                method = "GET";
                opr = item.get;
            }
            else if (item.post != null)
            {
                method = "POST";
                opr = item.post;
            }
            else if (item.put != null)
            {
                method = "PUT";
                opr = item.put;
            }
            else if (item.delete != null)
            {
                method = "DELETE";
                opr = item.delete;
            }
            else if (item.patch != null)
            {
                method = "PATCH";
                opr = item.patch;
            }
            else
            {
                method = "GET";
            }
            return opr;
        }

        private static List<string> CKeys = new List<string>();

        private static Dictionary<string, dynamic> InitParamsDefinedCache(IDictionary<string, Schema> item, string defkey, Dictionary<string, dynamic> dictResult, IDictionary<string, Schema> allitem)
        {
            if (dictResult == null)
            {
                dictResult = new Dictionary<string, dynamic>();
            }
            else if (dictResult.ContainsKey(defkey))
            {
                return dictResult[defkey];
            }

            if (CKeys.Contains(defkey))
            {
                return null;
            }
            CKeys.Add(defkey);
            if (defkey == "#/definitions/Citms.Utility.ApiResult[System.Collections.Generic.Dictionary[System.String,Citms.PIS.Model.Capture.CaptureTemplateDetail]]")
            {

            }
            Dictionary<string, dynamic> dictItem = new Dictionary<string, dynamic>();
            if (item != null && item.Keys.Count > 0)
            {
                foreach (var key in item.Keys)
                {
                    var prop = item[key];
                    if (prop.refDef == "#/definitions/System.Object")
                    {
                        dictItem[key] = new object();
                    }
                    else if (prop.type == "object")
                    {
                        if (prop.items == null)
                        {
                            continue;
                        }
                        dynamic r = null;
                        if (dictResult.TryGetValue(prop.items.refDef, out r))
                        {
                            dictItem[key] = r;
                        }
                        else
                        {
                            dictItem[key] = InitParamsDefinedCache(allitem[prop.items.refDef.Substring("#/definitions/".Length)].properties, prop.items.refDef, dictResult, allitem);
                        }
                    }
                    else if (prop.type == "array")
                    {
                        var list = new List<dynamic>();
                        if (prop.items == null)
                        {
                            continue;
                        }
                        if (!string.IsNullOrEmpty(prop.items.type))
                        {
                            if (prop.items.type == "string")
                            {
                                if (prop.items.format == "date-time")
                                {
                                    list.Add(DateTime.Now);
                                }
                                else
                                {
                                    list.Add("string");
                                }
                            }
                            else if (prop.items.type == "boolean")
                            {
                                list.Add(true);
                            }
                            else if (prop.items.type == "integer")
                            {
                                list.Add(0);
                            }
                            else if (prop.items.type == "number")
                            {
                                list.Add(0.0);
                            }
                        }
                        else
                        {
                            dynamic r = null;
                            if (dictResult.TryGetValue(prop.items.refDef, out r))
                            {
                                dictItem[key] = r;
                            }
                            else
                            {
                                list.Add(InitParamsDefinedCache(allitem[prop.items.refDef.Substring("#/definitions/".Length)].properties, prop.items.refDef, dictResult, allitem));
                            }

                        }
                        dictItem[key] = list;
                    }
                    else if (prop.type == "string")
                    {
                        if (prop.format == "date-time")
                        {
                            dictItem[key] = DateTime.Now;
                        }
                        else
                        {
                            dictItem[key] = "string";
                        }
                    }
                    else if (prop.type == "boolean")
                    {
                        dictItem[key] = true;
                    }
                    else if (prop.type == "integer")
                    {
                        dictItem[key] = 0;
                    }
                    else if (prop.type == "number")
                    {
                        dictItem[key] = 0.0;
                    }
                    else
                    {
                        //Console.WriteLine(prop.type);
                    }
                }
            }
            dictResult[defkey] = dictItem;
            return dictItem;
        }
    }
}
