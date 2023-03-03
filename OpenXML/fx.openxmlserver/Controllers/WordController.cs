using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System;
using System.IO;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.StaticFiles;
using FX.OpenXmlServer.WordHelpers;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using System.Linq;

namespace FX.OpenXmlServer.Controllers
{


    [Route("api/officemake")]
    public class WordController : Controller
    {
        [HttpGet("a")]///api/officemake/t
        public string a()
        {
            ExcelHelper wordHelper = new ExcelHelper();
            string str = "{\"EntrustUnitName\":\"深圳市生态环境局宝安管理局西乡所\"}";
            wordHelper.Init("C:\\Users\\Lee\\Desktop\\工作簿1.xlsx", "C:\\Users\\Lee\\Desktop\\xxx.xlsx", str);
            wordHelper.ExcelProcess();
            wordHelper.Close();

            /*            var dic = new Dictionary<string, string>();
                        dic.Add("ReportSignature1", @"D:\ServerAttachments\E-Signature\CMA清晰.png");
                        dic.Add("ReportSignature2", @"‪D:\ServerAttachments\E-Signature\CNAS清晰.png");
                        dic.Add("ReportSignature3", @"D:\ServerAttachments\E-Signature\检验检测章清晰.png");
                        dic.Add("OtherDiagram1", @"D:\ServerAttachments\E-Signature\检验检测章清晰.png");

                        WordHelper wordHelper = new WordHelper();
                        wordHelper.Init("D:\\ServerAttachments\\ReportTemplate\\BA_NoiseNoBackground.docx", "D:\\ServerAttachments\\Report\\2021\\2\\深圳市宝安区石岩隆兴抛光厂（测试）2021年02月03日111355510.docx");
                        wordHelper.WordProcess("{\"ReportSignature1\":\"ReportSignature1\",\"ReportSignature2\":\"ReportSignature2\",\"ReportSignature3\":\"ReportSignature3\",\"OtherDiagram\":[\"OtherDiagram1\"],\"ReportWord\":\"WZJ20210080\",\"ReportDate\":\"2021年02月03日\",\"MonitorUnitName\":\"深圳市宝安区石岩隆兴抛光厂（测试）\",\"MonitorUnitAddress\":\"深圳市生态环境监测站宝安分站\",\"EnvConditions\":\"风速为 0.6 m/s 晴\",\"PickTime\":\"2021-02-02 09:39-09:44\",\"InspectType\":\"监督监测\",\"MainSource\":\"生活噪声\",\"PickStaff\":\"管理员、罗子悦\",\"Instrument\":\"\",\"PickMethodStr\":\"GB 3096-2008\",\"ResultDataTable\":[{\"PlaceNo\":\"1 \",\"Place\":\"噪声#1 \",\"Leq\":\"52.3 \",\"BackGroundLeq\":\"60.1 \",\"CorrectionLeq\":\"* \"}],\"ResultDataTable2\":[]}", dic);
                        wordHelper.Close();*/

            /*            var serverAttachments = "C:\\ServerAttachments\\ZJ\\ReportTemplate\\ZJ_GeneralSubmissionTemplateExcel.xlsx";
                        var serverSaveReports = "C:\\ServerAttachments\\ZJ\\xx.xlsx";
                        var list = new List<List<String>>();
                        for (int i = 0; i < 20; i++)
                        {
                            var ls = new List<String> {
                                        i + "水温", "pH值", "溶解氧", "高锰酸盐指数", "化学需氧量", "五日生化需氧量", "氨氮",
                                        "总磷", "总氮", "铜", "锌", "氟化物", "硒", "砷", "汞", "镉", "六价铬", "铅", "氰化物",
                                        "挥发酚", "石油类", "阴离子表面活性剂", "硫化物", "粪大肠菌群", "硫酸盐","z","aa","ab"
                                    };
                            list.Add(ls);
                        }
                        //新  excelData
                        ExcelHelper eh = new ExcelHelper();
                        eh.Init(serverAttachments, serverSaveReports, "{}");
                        int r = 2;
                        foreach (var ls in list)//excelData)
                        {
                            eh.SetArray("Sheet1", "A", ++r, ls);
                        }

                        eh.Close();*/
            return "OK";


        }
        [HttpGet("t")]///api/officemake/t
        public string t() {
            string sourceFilePath = "C:\\Users\\Administrator\\Desktop\\t.docx";
            FileStream fsCurrent = new FileStream(sourceFilePath, FileMode.Open);
            WordprocessingDocument WordDocument = WordprocessingDocument.Open(fsCurrent, true);
            List<Table> IEnume = WordDocument.MainDocumentPart.Document.Body.Descendants<Table>().ToList();
            Table table = IEnume[0];

            var helper = new WordHelper();
            helper.AutoVerticalMerge(table,7);
            helper.AutoHorizontalMerge(table);
            WordDocument.MainDocumentPart.Document.Save();


            try

            {
                WordDocument.Close();
                WordDocument.Dispose();
                fsCurrent.Close();
            }
            catch (Exception e)
            {
                throw e;
            }



            return string.Empty;
        }




        /// <summary>
        /// 文档生成
        /// </summary>
        /// <param name="template">上传的模板文件</param>
        /// <param name="data">JSON字符串</param>
        /// <returns></returns>
        [HttpPost("DocxGenarate")]
        public object DocxGenarate(IFormFile template, [FromForm] string data)
        {
            try
            {
                IFormFileCollection Files = HttpContext.Request.Form.Files;
                Dictionary<string, string> dic = new Dictionary<string, string>();
                string CurrentPath = System.Environment.CurrentDirectory;
                string parentPath = Path.Combine(CurrentPath, "tmp");
                if (!Directory.Exists(parentPath)) Directory.CreateDirectory(parentPath);
                foreach (var file in Files)
                {
                    string guid = System.Guid.NewGuid().ToString();
                    string filePathName = Path.Combine(parentPath, guid + "_" + file.FileName);
                    Stream stream1 = file.OpenReadStream();
                    byte[] bytes = new byte[stream1.Length];
                    stream1.Read(bytes, 0, bytes.Length);
                    stream1.Seek(0, SeekOrigin.Begin);
                    FileStream fs = new FileStream(filePathName, FileMode.Create);
                    BinaryWriter bw = new BinaryWriter(fs);
                    bw.Write(bytes);
                    bw.Close();
                    fs.Close();
                    dic.Add(file.Name, filePathName);
                }

                string filePath = dic["template"];
                string fileExt = Path.GetExtension(template.FileName);
                WordHelper wordHelper = new WordHelper();
                wordHelper.Init(filePath);
                wordHelper.WordProcess(data, dic);
                wordHelper.Close();

                var stream = System.IO.File.OpenRead(filePath);

                //获取文件的ContentType
                var provider = new FileExtensionContentTypeProvider();
                var memi = provider.Mappings[fileExt];
                return File(stream, memi, Path.GetFileName(filePath));
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }


        /// <summary>
        /// 多个word文件按顺序整合
        /// </summary>
        /// <returns></returns>
        [HttpPost("MergeFiles")]
        public object MergeFiles() {
            try
            {
                IFormFileCollection Files = HttpContext.Request.Form.Files;
                List<string> sources = new List<string>();
                string CurrentPath = System.Environment.CurrentDirectory;
                string parentPath = Path.Combine(CurrentPath, "tmp");
                if (!Directory.Exists(parentPath)) Directory.CreateDirectory(parentPath);
                foreach (var file in Files)
                {
                    string guid = System.Guid.NewGuid().ToString();
                    string filePathName = Path.Combine(parentPath, guid + "_" + file.FileName);
                    Stream stream1 = file.OpenReadStream();
                    byte[] bytes = new byte[stream1.Length];
                    stream1.Read(bytes, 0, bytes.Length);
                    stream1.Seek(0, SeekOrigin.Begin);
                    FileStream fs = new FileStream(filePathName, FileMode.Create);
                    BinaryWriter bw = new BinaryWriter(fs);
                    bw.Write(bytes);
                    bw.Close();
                    fs.Close();
                    sources.Add(filePathName);
                }

                WordHelper wordHelper = new WordHelper();
                string mergeFile = wordHelper.MergeFiles(sources);
                var stream = System.IO.File.OpenRead(mergeFile);
                string fileExt = Path.GetExtension(mergeFile);

                //获取文件的ContentType
                var provider = new FileExtensionContentTypeProvider();
                var memi = provider.Mappings[fileExt];
                return File(stream, memi, Path.GetFileName(mergeFile));
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }
    }
}