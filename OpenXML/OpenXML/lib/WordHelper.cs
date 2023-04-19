using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;

namespace OpenXmlServer
{
    /// <summary>
    /// 批注实体类
    /// </summary>
    public class Comment
    {
        //批注选中的Key字符串
        public string Key;
        //批注唯一ID
        public string Id;
        //批注起始节点
        public OpenXmlElement Node;
        //批注文本内容指定的样式
        public Config config;

    }

    /// <summary>
    /// 书签实体类
    /// </summary>
    public class BookMark
    {
        //书签名
        public string Name;
        //书签唯一ID
        public string Id;
        //书签起始节点
        public BookmarkStart bookmarkStart;
    }

    /// <summary>
    /// 批注文本内容指定的样式的实体类
    /// </summary>
    public class Config
    {

        //数据类型
        public string DataType { get; set; }

        //字体大小
        public double? FontSize { get; set; }

        //字体
        public string FontFamily { get; set; }

        //（文字、段落的）对齐方式：0左，2居中，3右
        public int? JustificationValue { get; set; }

        //行距 一磅为20
        public int? SpacingBetweenLines { get; set; }

        //首行缩进多少字符 数值200，表示2字符
        public int? FirstLineChars { get; set; }

        //设置文本下划线
        public int? UnderlineValue { get; set; }

        //图片缩放比 1-100
        public int? ZoomRate { get; set; }

        //浮动图片水平位置，取值可是left, center, right或者数值0-3000000左右之间，数值越大越靠右
        public string HorizontalPosition { get; set; }

        //浮动图片垂直位置，可取数值0到8000000左右之间，数值越大越靠下
        public string VerticalPosition { get; set; }

        //表格纵向合并，取值为数组，如[0,1]，表示表格第一列和第二列同值合并
        public string VerticalMerge { get; set; }

        //表格横向合并，取值为true，表示真个表格都进行横向同值合并，不合并则不需要指定该属性
        public string HorizontalMerge { get; set; }


    }


    public class WordHelper
    {

        /// <summary>
        /// 当前打开的word文档对象,操作完成后,请调用CloseDocument关闭文档,释放文件占用
        /// </summary>
        public WordprocessingDocument WordDocument { get; private set; }

        /// <summary>
        /// 打开word文档的文件流,操作完成后,请调用CloseDocument关闭文档,释放文件占用
        /// </summary>
        private FileStream fsCurrent = null;

        /// <summary>
        /// 批注列表
        /// </summary>
        public List<Comment> CommentList = new List<Comment>();

        /// <summary>
        /// 书签列表
        /// </summary>
        public List<BookMark> BookMarkList = new List<BookMark>();

        /// <summary>
        /// 打开文档及初始化
        /// </summary>
        /// <param name="filePath"></param>
        public void Init(string filePath, string destinationFilePath = null)
        {
            OpenDocument(filePath, destinationFilePath);
            InitCommentList();
            InitBookMarkList();
        }

        /// <summary>
        /// 打开文档
        /// </summary>
        /// <param name="sourceFilePath">源文档</param>
        /// <param name="destinationFilePath">目标文档</param>
        /// <returns></returns>
        public WordprocessingDocument OpenDocument(string sourceFilePath, string destinationFilePath = null)
        {
            if (WordDocument != null)
            {
                WordDocument.Close();
                WordDocument.Dispose();
            }
            if (fsCurrent != null)
            {
                fsCurrent.Close();
            }
            try
            {
                CheckIf07PlusDocx(sourceFilePath);

                if (!string.IsNullOrWhiteSpace(destinationFilePath))
                {
                    CheckIf07PlusDocx(destinationFilePath);

                    File.Copy(sourceFilePath, destinationFilePath, true);

                    fsCurrent = new FileStream(destinationFilePath, FileMode.Open); // 必须以Stream的默认形式打开文件 
                }
                else
                {
                    fsCurrent = new FileStream(sourceFilePath, FileMode.Open); // 必须以Stream的默认形式打开文件 
                }

                WordDocument = WordprocessingDocument.Open(fsCurrent, true);
                return WordDocument;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 获取批注选中文字及批注
        /// </summary>
        public void InitCommentList()
        {
            string KeyStr = "";
            string get(OpenXmlElement item)
            {
                if (item.NextSibling().LocalName != "commentRangeEnd")
                {
                    KeyStr += item.NextSibling().InnerText;
                    get(item.NextSibling());
                }
                return KeyStr;
            }
            var CommentRangeStarts = WordDocument.MainDocumentPart.Document.Descendants<CommentRangeStart>();


            foreach (var item in CommentRangeStarts)
            {
                string Key = get(item);
                Comment comment = new Comment();
                comment.Key = Key;
                comment.Id = item.Id;
                comment.Node = item;
                CommentList.Add(comment);
                KeyStr = "";
            }
            var commentsParts = WordDocument.MainDocumentPart.WordprocessingCommentsPart;
            if (commentsParts == null) return;
            var comments = commentsParts.Comments.Descendants<DocumentFormat.OpenXml.Wordprocessing.Comment>();
            foreach (var comment in comments)
            {
                foreach (Comment Comment in CommentList)
                {
                    if (comment.Id == Comment.Id)
                    {
                        Comment.config = GetConfig(comment.InnerText);
                    }
                }
            }
        }

        /// <summary>
        /// 初始化书签列表
        /// </summary>
        public void InitBookMarkList()
        {
            //页眉书签
            foreach (HeaderPart headerPart in WordDocument.MainDocumentPart.HeaderParts)
            {
                foreach (BookmarkStart bookmarkStart in headerPart.RootElement.Descendants<BookmarkStart>())
                {
                    if (bookmarkStart.Name.ToString() == "_GoBack") continue;

                    BookMark bookMark = new BookMark();
                    bookMark.Name = bookmarkStart.Name.ToString().Trim();
                    bookMark.Id = bookmarkStart.Id.ToString();
                    bookMark.bookmarkStart = bookmarkStart;
                    BookMarkList.Add(bookMark);
                }
            }

            //页面书签
            foreach (BookmarkStart bookmarkStart in WordDocument.MainDocumentPart.Document.Descendants<BookmarkStart>())
            {
                if (bookmarkStart.Name.ToString() == "_GoBack") continue;

                BookMark bookMark = new BookMark();
                bookMark.Name = bookmarkStart.Name.ToString().Trim();
                bookMark.Id = bookmarkStart.Id.ToString();
                bookMark.bookmarkStart = bookmarkStart;
                BookMarkList.Add(bookMark);
            }

            //页脚书签
            foreach (FooterPart footerPart in WordDocument.MainDocumentPart.FooterParts)
            {
                foreach (BookmarkStart bookmarkStart in footerPart.RootElement.Descendants<BookmarkStart>())
                {
                    if (bookmarkStart.Name.ToString() == "_GoBack") continue;

                    BookMark bookMark = new BookMark();
                    bookMark.Name = bookmarkStart.Name.ToString().Trim();
                    bookMark.Id = bookmarkStart.Id.ToString();
                    bookMark.bookmarkStart = bookmarkStart;
                    BookMarkList.Add(bookMark);
                }
            }


        }


        /// <summary>
        /// 文字处理生成文档
        /// </summary>
        /// <param name="json">用户提交的json数据</param>
        /// <param name="dic">图片字段及图片路径字典数据</param>
        public void WordProcess(string json, Dictionary<string, string> dic)
        {
            try
            {
                JObject item = (JObject)JsonConvert.DeserializeObject(json);

                //处理批注
                foreach (Comment comment in CommentList)
                {
                    Config config = comment.config;

                    if (config.DataType == null) throw new Exception("批注没有数据类型字段");
                    Object data = GetData(item, comment.Key, config.DataType);

                    switch (config.DataType)
                    {
                        case ("0"):
                            Run run = (Run)CreateRunText((string)data, config);
                            if (run.RunProperties == null)
                            {
                                Run r = (Run)comment.Node.NextSibling<Run>();
                                run.RunProperties = r.RunProperties != null ? r.RunProperties.Clone() as RunProperties : null;
                            }
                            comment.Node.InsertBeforeSelf(run);
                            break;
                        case ("1"):
                            if (data is JArray)
                            {
                                foreach (string imgKey in ((JArray)data))
                                {
                                    string url = dic.ContainsKey(imgKey) ? dic[imgKey] : string.Empty;
                                    if (!System.IO.File.Exists(url))
                                    {
                                        DeleteCommentSelectedContent(comment.Node);
                                        continue;
                                    }
                                    Paragraph paragraph = new Paragraph();
                                    Paragraph perantParagraph = comment.Node.Ancestors<Paragraph>().FirstOrDefault();
                                    paragraph.ParagraphProperties = perantParagraph.ParagraphProperties.Clone() as ParagraphProperties;
                                    paragraph.Append(CreateImage(url, config));
                                    perantParagraph.InsertAfterSelf(paragraph);
                                }
                            }
                            else
                            {
                                string url = dic.ContainsKey((string)data) ? dic[(string)data] : string.Empty;
                                if (!System.IO.File.Exists(url))
                                {
                                    DeleteCommentSelectedContent(comment.Node);
                                    continue;
                                }
                                Paragraph paragraph = new Paragraph();
                                Paragraph perantParagraph = comment.Node.Ancestors<Paragraph>().FirstOrDefault();
                                //paragraph.ParagraphProperties = perantParagraph.ParagraphProperties.Clone() as ParagraphProperties;
                                //paragraph.Append(CreateImage(url, config));
                                //perantParagraph.InsertAfterSelf(paragraph);
                                perantParagraph.Append(CreateImage(url, config));
                            }
                            break;
                        case ("2"):
                            Run checkBox = (Run)CreateCheckBox((bool)data, config);
                            if (checkBox.RunProperties == null)
                            {
                                Run r = (Run)comment.Node.NextSibling<Run>();
                                checkBox.RunProperties = r.RunProperties != null ? r.RunProperties.Clone() as RunProperties : null;
                            }

                            comment.Node.InsertBeforeSelf(checkBox);
                            break;
                        case ("3"):
                            OpenXmlElement newParagraph = CreateParagraph((string)data, config);
                            Paragraph oldParagraph = comment.Node.Ancestors<Paragraph>().FirstOrDefault();
                            oldParagraph.InsertAfterSelf(newParagraph);
                            oldParagraph.Remove();
                            break;
                        case ("4"):
                            OpenXmlElement table = CreateTable((JObject)data, config, dic);
                            comment.Node.Parent.InsertAfterSelf(table);
                            break;
                        case ("5"):
                            FillTable((JArray)data, config, comment.Node);
                            break;
                        default:
                            break;
                    }
                    DeleteCommentSelectedContent(comment.Node);
                }

                DeleteComments();


                //处理书签
                foreach (BookMark bookMark in BookMarkList) 
                {
                    if (item.ContainsKey(bookMark.Name)) 
                    {
                        string data = (string)item[bookMark.Name];
                        Run run;
                        if (File.Exists(data))
                            run = (Run)CreateImage(data, new Config() { ZoomRate = 30 });
                        else
                            run = (Run)CreateRunText(data, new Config());

                        bookMark.bookmarkStart.InsertAfterSelf(run);
                        bookMark.bookmarkStart.NextSibling<BookmarkEnd>().Remove();
                        bookMark.bookmarkStart.Remove();
                    }
                }

                WordDocument.MainDocumentPart.Document.Save();

            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /// <summary>
        /// 根据批注选中文本从JSON数据中过去对应的数据
        /// </summary>
        /// <param name="item">JSON数据对象</param>
        /// <param name="Key">批注选中文本表示的Key</param>
        /// <param name="DataType">数据类型</param>
        /// <returns></returns>
        public Object GetData(JObject item, string Key, string DataType)
        {
            try
            {
                string[] KeyArr = Key.Split('.');
                string MsgKey = "";
                Object val;
                int i = 0;
                while (true)
                {
                    MsgKey += "." + KeyArr[i];
                    if (item.ContainsKey(KeyArr[i]))
                    {
                        if ((i + 1) == KeyArr.Length)
                        {
                            switch (DataType) {
                                case "1":
                                    //val = item[KeyArr[i]] is JArray ? (JArray)item[KeyArr[i]] : (string)item[KeyArr[i]];
                                    if (item[KeyArr[i]] is JArray)
                                        val = (JArray)item[KeyArr[i]];
                                    else
                                        val = (string)item[KeyArr[i]];
                                    break;
                                case "2":
                                    val = (bool)item[KeyArr[i]];
                                    break;
                                case "4":
                                    val = (JObject)item[KeyArr[i]];
                                    break;
                                case "5":
                                    val = (JArray)item[KeyArr[i]];
                                    break;
                                default:
                                    val = (string)item[KeyArr[i]];
                                    break;
                            }
                            break;
                        }
                        else
                        {
                            if (item[KeyArr[i]] is JObject)
                                item = (JObject)item[KeyArr[i]];
                            else
                                throw new Exception("JSON数据异常");
                            i++;
                        }
                    }
                    else
                    {
                        throw new Exception("【" + MsgKey.Substring(1) + "】字段不存在");
                    }
                }
                return val;
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        /// <summary>
        /// 获取批注文本内容指定的样式
        /// </summary>
        /// <param name="rawCfg">批注里面的指定样式的文本内容</param>
        /// <returns></returns>
        public Config GetConfig(string rawCfg)
        {
            Config config = new Config();
            string[] Properties = rawCfg.Replace(" ", string.Empty).Split('|', '\n');
            foreach (string property in Properties)
            {
                string[] arr = property.Split('=');
                switch (arr[0])
                {
                    case ("DataType"):
                        config.DataType = arr[1];
                        break;
                    case ("FontSize"):
                        config.FontSize = int.Parse(arr[1]);
                        break;
                    case ("FontFamily"):
                        config.FontFamily = arr[1];
                        break;
                    case ("JustificationValue"):
                        config.JustificationValue = int.Parse(arr[1]);
                        break;
                    case ("SpacingBetweenLines"):
                        config.SpacingBetweenLines = int.Parse(arr[1]);
                        break;
                    case ("FirstLineChars"):
                        config.FirstLineChars = int.Parse(arr[1]);
                        break;
                    case ("UnderlineValue"):
                        config.UnderlineValue = int.Parse(arr[1]);
                        break;
                    case ("ZoomRate"):
                        config.ZoomRate = int.Parse(arr[1]);
                        break;
                    case ("HorizontalPosition"):
                        config.HorizontalPosition = arr[1];
                        break;
                    case ("VerticalPosition"):
                        config.VerticalPosition = arr[1];
                        break;
                    case ("VerticalMerge"):
                        config.VerticalMerge = arr[1];
                        break;
                    case ("HorizontalMerge"):
                        config.HorizontalMerge = arr[1];
                        break;
                    default:
                        break;
                }
            }
            return config;

        }
        /// <summary>
        /// 删除所有批注信息
        /// </summary>
        public void DeleteComments()
        {
            var wordprocessingCommentsPart = WordDocument.MainDocumentPart.WordprocessingCommentsPart;
            if (wordprocessingCommentsPart == null) return;
            List<DocumentFormat.OpenXml.Wordprocessing.Comment> commentsToDelete = wordprocessingCommentsPart.Comments.Elements<DocumentFormat.OpenXml.Wordprocessing.Comment>().ToList();
            foreach (DocumentFormat.OpenXml.Wordprocessing.Comment c in commentsToDelete)
            {
                c.Remove();
            }

            IEnumerable<string> commentIds = commentsToDelete.Select(r => r.Id.Value);

            List<CommentRangeStart> commentRangeStartToDelete = WordDocument.MainDocumentPart.Document.Descendants<CommentRangeStart>().
            Where(c => commentIds.Contains(c.Id.Value)).ToList();
            foreach (CommentRangeStart c in commentRangeStartToDelete)
            {
                c.Remove();
            }

            List<CommentRangeEnd> commentRangeEndToDelete = WordDocument.MainDocumentPart.Document.Descendants<CommentRangeEnd>().Where(c => commentIds.Contains(c.Id.Value)).ToList();
            foreach (CommentRangeEnd c in commentRangeEndToDelete)
            {
                c.Remove();
            }

            List<CommentReference> commentRangeReferenceToDelete = WordDocument.MainDocumentPart.Document.Descendants<CommentReference>().Where(c => commentIds.Contains(c.Id.Value)).ToList();
            foreach (CommentReference c in commentRangeReferenceToDelete)
            {
                c.Remove();
            }
        }

        /// <summary>
        /// 递归删除批注选中的内容节点
        /// </summary>
        /// <param name="ele">批注开始节点</param>
        public void DeleteCommentSelectedContent(OpenXmlElement ele)
        {
            if (ele.NextSibling().LocalName != "commentRangeEnd")
            {
                DeleteCommentSelectedContent(ele.NextSibling());
                ele.NextSibling().Remove();
            }
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Close()
        {
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

        }



        /// <summary>
        /// 填充表格
        /// </summary>
        /// <param name="tb">JArray数组</param>
        /// <param name="config">批注指定的样式</param>
        /// <param name="node">批注开始节点</param>
        public void FillTable(JArray tb, Config config, OpenXmlElement node)
        {
            OpenXmlElement curRow = node.Ancestors<TableRow>().ToList()[0];

            for (int i = 0; i < tb.Count; i++)
            {
                JObject jobj = (JObject)tb[i];

                OpenXmlElement tableRow = (OpenXmlElement)curRow.Clone();
                OpenXmlElement curCell = tableRow.GetFirstChild<TableCell>();
                foreach (JProperty jp in jobj.Properties())
                {
                    Paragraph newParagraph = (Paragraph)CreateParagraph((string)jp.Value, config);
                    Paragraph oldParagraph = curCell.GetFirstChild<Paragraph>();
                    newParagraph.ParagraphProperties = oldParagraph.ParagraphProperties.Clone() as ParagraphProperties;

                    curCell.ReplaceChild<Paragraph>(newParagraph, oldParagraph);
                    curCell = curCell.NextSibling<TableCell>();
                }

                curRow.InsertBeforeSelf(tableRow);

            }
            Table table = (Table)curRow.Parent;
            curRow.Remove();

            JArray VerticalMerge = config.VerticalMerge != null ? (JArray)JsonConvert.DeserializeObject(config.VerticalMerge) : null;
            string HorizontalMerge = config.HorizontalMerge != null ? config.HorizontalMerge : null;

            if (VerticalMerge != null && VerticalMerge.Count > 0)
            {
                foreach (int colIndex in VerticalMerge)
                {
                    AutoVerticalMerge(table, colIndex);
                }
            }

            if (HorizontalMerge != null && HorizontalMerge == "true")
            {
                AutoHorizontalMerge(table);
            }
        }

        /// <summary>
        /// 根据json串创建表格
        /// </summary>
        /// <param name="jobj">表格结构样式及数据的JSON对象</param>
        /// <param name="config">批注指定的样式</param>
        /// <param name="dic">图片字段及图片路径字典数据</param>
        /// <returns></returns>
        public OpenXmlElement CreateTable(JObject jobj, Config config, Dictionary<string, string> dic)
        {
            JArray columns = (JArray)jobj["columns"];
            JArray rows = (JArray)jobj["rows"];
            JArray cells = (JArray)jobj["cells"];
            if (rows.Count != cells.Count)
                throw new Exception("表格JSON数据格式错误");

            //表格JSON数据预处理
            JObject lastRowSpanDict = null;
            for (int r = 0; r < cells.Count; r++)
            {
                JArray rData = (JArray)cells[r];
                int lastcolspan = 0;
                JObject currentRowSpanDict = new JObject();
                for (int cIndex = 0; cIndex < rData.Count; cIndex++)
                {
                    JObject cData = (JObject)rData[cIndex];
                    int colspan = 1;
                    int rowspan = 1;
                    if (cData.ContainsKey("colspan")) colspan = (int)cData["colspan"];
                    if (cData.ContainsKey("rowspan")) rowspan = (int)cData["rowspan"];

                    int offset = lastcolspan;
                    JObject obj = new JObject();
                    if (lastRowSpanDict != null && lastRowSpanDict.ContainsKey(offset.ToString())
                        && (int)lastRowSpanDict[offset.ToString()]["colspan"] == colspan
                        && (int)lastRowSpanDict[offset.ToString()]["leftrowspan"] > 0)
                    {
                        cData.Add("verticalMerge", 0);
                        ((JObject)lastRowSpanDict[offset.ToString()])["leftrowspan"] = (int)((JObject)lastRowSpanDict[offset.ToString()])["leftrowspan"] - 1;


                        obj.Add("colspan", colspan);
                        obj.Add("leftrowspan", (int)((JObject)lastRowSpanDict[offset.ToString()])["leftrowspan"]);
                        obj.Add("verticalMerge", 0);

                        currentRowSpanDict.Add(offset.ToString(), obj);
                    }
                    else
                    {
                        obj.Add("colspan", colspan);
                        obj.Add("leftrowspan", rowspan - 1);
                        obj.Add("verticalMerge", rowspan > 1 ? 1 : 0);

                        currentRowSpanDict.Add(offset.ToString(), obj);

                        if (rowspan > 1) cData["verticalMerge"] = 1;
                    }
                    lastcolspan += colspan;
                }

                lastRowSpanDict = currentRowSpanDict;
                if (lastcolspan > columns.Count) throw new Exception();
            }



            //开始根据JSON数据创建表格
            Table table = new Table();

            //累加所有列宽作为表宽
            int widthCount = 0;

            //每个列宽存为数组，后面好确定每个单元格宽度
            List<int> columnWidths = new List<int>();

            TableGrid tableGrid = new TableGrid();
            for (int i = 0; i < columns.Count; i++)
            {
                JObject obj = (JObject)columns[i];
                GridColumn gridColumn = new GridColumn() { Width = obj["width"].ToString() };
                tableGrid.Append(gridColumn);

                widthCount += (int)obj["width"];
                columnWidths.Add((int)obj["width"]);
            }

            TableProperties tableProperties = new TableProperties(new TableBorders(
                new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
            ));
            TableStyle tableStyle = new TableStyle() { Val = "a4" };
            TableWidth tableWidth = new TableWidth() { Width = widthCount.ToString(), Type = TableWidthUnitValues.Dxa };
            TableLook tableLook = new TableLook() { Val = "04A0" };

            tableProperties.Append(tableStyle);
            tableProperties.Append(tableWidth);
            tableProperties.Append(tableLook);

            table.Append(tableProperties);
            table.Append(tableGrid);


            for (int i = 0; i < cells.Count; i++)
            {
                JArray rowCells = (JArray)cells[i];
                TableRow tableRow = new TableRow() { RsidTableRowAddition = "00E858D7", RsidTableRowProperties = "00E858D7" };
                TableRowProperties tableRowProperties = new TableRowProperties();
                if (((JObject)rows[i]).ContainsKey("height"))
                {
                    uint height = (uint)((JObject)rows[i])["height"];
                    TableRowHeight tableRowHeight = new TableRowHeight() { Val = (UInt32Value)height };
                    tableRowProperties.Append(tableRowHeight);
                }
                tableRow.Append(tableRowProperties);


                List<int> cellWidths = new List<int>(columnWidths.ToArray());
                for (int j = 0; j < rowCells.Count; j++)
                {
                    JObject cell = (JObject)rowCells[j];

                    //根据列宽列表，计算每个单元格的宽度
                    int colspan = cell.ContainsKey("colspan") ? (int)cell["colspan"] : 1;
                    int cellWidth = 0;
                    for (int k = 0; k < colspan; k ++) {
                        cellWidth += cellWidths.First();
                        cellWidths.RemoveAt(0);
                    }

                    //开始创建单元格
                    TableCell tableCell = new TableCell();
                    TableCellProperties tableCellProperties = new TableCellProperties();
                    TableCellWidth tableCellWidth = new TableCellWidth() { Width = cellWidth.ToString(), Type = TableWidthUnitValues.Dxa };
                    TableCellVerticalAlignment tableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                    if (cell.ContainsKey("colspan"))
                    {
                        GridSpan gridSpan = new GridSpan() { Val = (int)cell["colspan"] };
                        tableCellProperties.Append(gridSpan);
                    }
                    if (cell.ContainsKey("verticalMerge") && (int)cell["verticalMerge"] == 1)
                    {
                        VerticalMerge verticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
                        tableCellProperties.Append(verticalMerge);
                    }
                    else if (cell.ContainsKey("verticalMerge") && (int)cell["verticalMerge"] == 0)
                    {
                        VerticalMerge verticalMerge = new VerticalMerge();
                        tableCellProperties.Append(verticalMerge);
                    }
                    tableCellProperties.Append(tableCellWidth);
                    tableCellProperties.Append(tableCellVerticalAlignment);
                    tableCell.Append(tableCellProperties);

                    Paragraph paragraph;
                    if (cell.ContainsKey("DataType") && (int)cell["DataType"] == 1)
                    {
                        if (!dic.ContainsKey((string)cell["data"]))
                            throw new Exception("图片【" + (string)cell["data"] + "】未上传");
                        string url = dic[(string)cell["data"]];
                        if (!System.IO.File.Exists(url))
                            throw new Exception("图片【" + (string)cell["data"] + "】不存在");
                        paragraph = new Paragraph();
                        ParagraphProperties paragraphProperties = new ParagraphProperties();
                        int? JustificationValue = config.JustificationValue != null ? config.JustificationValue : 2;
                        paragraph.Append(CreateImage(url, config));
                        paragraphProperties.AppendChild(new Justification() { Val = (JustificationValues)JustificationValue });
                        paragraph.ParagraphProperties = paragraphProperties;
                    }
                    else
                    {
                        paragraph = (Paragraph)CreateParagraph((string)cell["data"], config);
                    }
                    tableCell.Append(paragraph);
                    tableRow.Append(tableCell);
                }
                table.Append(tableRow);
            }
            return table;
        }

        /// <summary>
        /// 创建图片
        /// </summary>
        /// <param name="url">图片地址</param>
        /// <param name="config">批注指定的样式</param>
        /// <returns></returns>
        public OpenXmlElement CreateImage(string url, Config config)
        {
            var imageType = ImagePartType.Jpeg;

            #region other image type

            var extension = Path.GetExtension(url);
            if (extension != null)
            {
                var ext = extension.TrimStart('.').ToLower();
                switch (ext)
                {
                    case "jpg":
                    case "jpeg":
                    case "jpe":
                        imageType = ImagePartType.Jpeg;
                        break;
                    case "bmp":
                        imageType = ImagePartType.Bmp;
                        break;
                    case "gif":
                        imageType = ImagePartType.Gif;
                        break;
                    case "icon":
                        imageType = ImagePartType.Icon;
                        break;
                    case "png":
                        imageType = ImagePartType.Png;
                        break;
                    case "tiff":
                        imageType = ImagePartType.Tiff;
                        break;
                }
            }
            #endregion

            string relationshipId;
            long imageWidthEmu = 0; long imageHeightEmu = 0;
            using (FileStream fsImageFile = new FileStream(url, FileMode.Open, FileAccess.Read))
            {
                var maxWidthCm = 16.51;
                const int emusPerInch = 914400;
                const int emusPerCm = 360000;
                //var imageFile = Image.FromFile(url);
                //imageWidthEmu = (long)((imageFile.Width / imageFile.HorizontalResolution) * emusPerInch);
                //imageHeightEmu = (long)((imageFile.Height / imageFile.VerticalResolution) * emusPerInch);

                var maxWidthEmus = (long)(maxWidthCm * emusPerCm);
                if (config.ZoomRate > 0)
                {
                    imageWidthEmu = (long)config.ZoomRate * imageWidthEmu / 100;
                    imageHeightEmu = (long)config.ZoomRate * imageHeightEmu / 100;
                }
                else if (imageWidthEmu > maxWidthEmus)
                {
                    //超出最大宽度，强制压缩，跟word一样的
                    var ratio = (imageHeightEmu * 1.0m) / imageWidthEmu;
                    imageWidthEmu = maxWidthEmus;
                    imageHeightEmu = (long)(imageWidthEmu * ratio);
                }

                ImagePart imagePart = WordDocument.MainDocumentPart.AddImagePart(imageType);
                imagePart.FeedData(fsImageFile);

                relationshipId = WordDocument.MainDocumentPart.GetIdOfPart(imagePart);
            }


            if (config.HorizontalPosition != null && config.VerticalPosition != null)
            {
                string[] alignments = { "left", "center", "right" };
                IList<string> alignmentsList = (IList<string>)alignments;
                Object position = new Object();
                if (alignmentsList.Contains(config.HorizontalPosition))
                    position = new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignment() { Text = config.HorizontalPosition };
                else
                    position = new DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset() { Text = config.HorizontalPosition };


                DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties = new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties();
                DocumentFormat.OpenXml.Drawing.GraphicFrameLocks graphicFrameLocks = new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true };
                graphicFrameLocks.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                nonVisualGraphicFrameDrawingProperties.Append(graphicFrameLocks);

                DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi useLocalDpi = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi() { Val = false };
                useLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");


                DocumentFormat.OpenXml.Drawing.Pictures.Picture picture = new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "img1.png" },
                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
                        ),
                        new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                            new DocumentFormat.OpenXml.Drawing.Blip(
                                new DocumentFormat.OpenXml.Drawing.BlipExtensionList(
                                new DocumentFormat.OpenXml.Drawing.BlipExtension(
                                    useLocalDpi
                                )
                                { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }
                            ))
                            { Embed = relationshipId, CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print },
                            new DocumentFormat.OpenXml.Drawing.Stretch(
                                new DocumentFormat.OpenXml.Drawing.FillRectangle()
                            )
                        ),
                        new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                            new DocumentFormat.OpenXml.Drawing.Transform2D(
                            new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                            new DocumentFormat.OpenXml.Drawing.Extents() { Cx = imageWidthEmu, Cy = imageHeightEmu }
                            ),
                            new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                            new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                            )
                            { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }
                        )
                    );
                picture.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");



                DocumentFormat.OpenXml.Drawing.Graphic graphic = new DocumentFormat.OpenXml.Drawing.Graphic(
                    new DocumentFormat.OpenXml.Drawing.GraphicData(picture) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    );
                graphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                Drawing anchorDrawing = new Drawing(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor(
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.SimplePosition() { X = 0L, Y = 0L },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalPosition(
                            (OpenXmlElement)position
                        )
                        { RelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Margin },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalPosition(
                            new DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset() { Text = config.VerticalPosition }
                        )
                        { RelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Paragraph },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = imageWidthEmu, Cy = imageHeightEmu },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapNone(),
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = (UInt32Value)1U, Name = "图片 1" },
                        nonVisualGraphicFrameDrawingProperties,
                        graphic)
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)114300U,
                        DistanceFromRight = (UInt32Value)114300U,
                        SimplePos = false,
                        RelativeHeight = (UInt32Value)251658240U,
                        BehindDoc = true,
                        Locked = false,
                        LayoutInCell = true,
                        AllowOverlap = true,
                        EditId = "60CCEEB3",
                        AnchorId = "366877E9"
                    }
                    );

                RunProperties runProperties = new RunProperties();
                NoProof noProof = new NoProof();
                FontSize fontSize = new FontSize() { Val = "30" };
                FontSizeComplexScript fontSizeComplexScript = new FontSizeComplexScript() { Val = "30" };

                runProperties.Append(noProof);
                runProperties.Append(fontSize);
                runProperties.Append(fontSizeComplexScript);

                return new Run(runProperties, anchorDrawing);

            }

            Drawing drawing =
                 new Drawing(
                     new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = imageWidthEmu, Cy = imageHeightEmu }, //缩放100%
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }),
                         new DocumentFormat.OpenXml.Drawing.Graphic(
                             new DocumentFormat.OpenXml.Drawing.GraphicData(
                                 new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                     new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                         new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
                                         ),
                                     new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                         new DocumentFormat.OpenXml.Drawing.Blip(new DocumentFormat.OpenXml.Drawing.BlipExtensionList(new DocumentFormat.OpenXml.Drawing.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }))
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                                         },
                                         new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())
                                         ),
                                     new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                         new DocumentFormat.OpenXml.Drawing.Transform2D(
                                             new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                             new DocumentFormat.OpenXml.Drawing.Extents() { Cx = imageWidthEmu, Cy = imageHeightEmu }
                                             ),
                                         new DocumentFormat.OpenXml.Drawing.PresetGeometry(new DocumentFormat.OpenXml.Drawing.AdjustValueList()) { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }
                                         )
                                    )
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         //EditId = "50D07946" //设置了此值后Office2007会有问题 xzk 2016-12-04
                     });


            return new Run(drawing);
        }

        /// <summary>
        /// 创建一段文本
        /// </summary>
        /// <param name="text">文本内容</param>
        /// <param name="config">批注指定的样式</param>
        /// <returns></returns>
        public OpenXmlElement CreateRunText(string text, Config config)
        {
            var rPr = new RunProperties();

            if (!string.IsNullOrEmpty(config.FontFamily))
            {
                rPr.Append(new RunFonts() { Ascii = config.FontFamily, HighAnsi = config.FontFamily, EastAsia = config.FontFamily });
            }

            if (config.FontSize > 0)
            {
                rPr.Append(new FontSize() { Val = config.FontSize.ToString() });
            }
            OpenXmlElement oxe = rPr.HasChildren ? new Run(rPr, new Text(text)) : new Run(new Text(text));
            return oxe;
        }

        /// <summary>
        /// 创建复选框
        /// </summary>
        /// <param name="isChecked">打勾：true；空格：false</param>
        /// <param name="config">批注指定的样式</param>
        /// <returns></returns>
        public OpenXmlElement CreateCheckBox(bool isChecked, Config config)
        {
            var rPr = new RunProperties();

            if (!string.IsNullOrEmpty(config.FontFamily))
            {
                rPr.Append(new RunFonts() { Ascii = config.FontFamily, HighAnsi = config.FontFamily, EastAsia = config.FontFamily });
            }

            if (config.FontSize > 0)
            {
                rPr.Append(new FontSize() { Val = config.FontSize.ToString() });
            }
            OpenXmlElement oxe = rPr.HasChildren ?
                new Run(rPr, new SymbolChar() { Font = "Wingdings 2", Char = isChecked ? "F052" : "F0A3" }) :
                new Run(new SymbolChar() { Font = "Wingdings 2", Char = isChecked ? "F052" : "F0A3" });
            return oxe;
        }

        /// <summary>
        /// 页眉
        /// </summary>
        /// <param name="headerText"></param>
        /// <returns></returns>
        public Header CreatePageHeaderPart(string headerText)
        {
            // set the position to be the center  
            PositionalTab pTab = new PositionalTab()
            {
                Alignment = AbsolutePositionTabAlignmentValues.Center,
                RelativeTo = AbsolutePositionTabPositioningBaseValues.Margin,
                Leader = AbsolutePositionTabLeaderCharValues.None
            };

            var element =
              new Header(
                new Paragraph(
                  new ParagraphProperties(new ParagraphStyleId() { Val = "Header" }),
                  new Run(pTab, new Text(headerText))
                )
              );

            return element;
        }

        /// <summary>
        /// 创建段落
        /// </summary>
        /// <param name="text">文本信息</param>
        /// <param name="config">批注指定的样式</param>
        /// <returns></returns>
        public OpenXmlElement CreateParagraph(string text, Config config)
        {
            var pPr = new ParagraphProperties();
            if (config.JustificationValue != null)
            {
                pPr.AppendChild(new Justification() { Val = (JustificationValues)config.JustificationValue });
            }
            if (config.SpacingBetweenLines != null)
            {
                pPr.AppendChild(new SpacingBetweenLines() { Line = config.SpacingBetweenLines.ToString(), LineRule = LineSpacingRuleValues.Exact });
            }

            if (config.FirstLineChars > 0)
            {
                pPr.AppendChild(new Indentation() { FirstLineChars = config.FirstLineChars });
            }

            var rPrForPar = !string.IsNullOrWhiteSpace(config.FontFamily)
                ? new RunProperties(new RunFonts()
                {
                    Ascii = config.FontFamily,
                    HighAnsi = config.FontFamily,
                    EastAsia = config.FontFamily
                })
                : new RunProperties(new RunFonts());

            if (config.FontSize != null)
            {
                double FontSize_Pound = (double)config.FontSize * 2;
                rPrForPar.Append(new FontSize() { Val = FontSize_Pound.ToString() });
                rPrForPar.Append(new FontSizeComplexScript() { Val = FontSize_Pound.ToString() });
            }
            if (config.UnderlineValue != null)
            {
                rPrForPar.AppendChild(new Underline() { Val = (UnderlineValues)config.UnderlineValue });
            }

            pPr.AppendChild(rPrForPar);

            var rPrForOxe = !string.IsNullOrWhiteSpace(config.FontFamily)
                ? new RunProperties(new RunFonts()
                {
                    Ascii = config.FontFamily,
                    HighAnsi = config.FontFamily,
                    EastAsia = config.FontFamily
                })
                : new RunProperties();
            if (config.FontSize != null)
            {
                double FontSize_Pound = (double)config.FontSize * 2;
                rPrForOxe.Append(new FontSize() { Val = FontSize_Pound.ToString() });
            }

            OpenXmlElement oxe = new Paragraph(pPr);
            if (!string.IsNullOrEmpty(text))
            {
                Run run = new Run(rPrForOxe, new Text(text));
                oxe.AppendChild(run);
            }

            return oxe;

        }


        /// <summary>
        /// 表格自动纵向合并
        /// </summary>
        /// <param name="table">表格</param>
        /// <param name="cIndex">需要合并的列</param>
        public void AutoVerticalMerge(Table table, int cIndex)
        {
            List<TableRow> tableRows = table.Descendants<TableRow>().ToList();
            for (int i = 0; i < tableRows.Count; i++)
            {
                List<TableCell> rowCells = tableRows[i].Descendants<TableCell>().ToList();
                if (rowCells.Count <= cIndex) continue;
                TableCell curCell = rowCells[cIndex];
                int nextIndex = i + 1;
                if (nextIndex == tableRows.Count) break;
                List<TableCell> nextRowCells = tableRows[nextIndex].Descendants<TableCell>().ToList();
                if (nextRowCells.Count <= cIndex) continue;
                TableCell nextCell = nextRowCells[cIndex];

                int curCellGridSpan = curCell.TableCellProperties.GridSpan != null ? curCell.TableCellProperties.GridSpan.Val.Value : 0;
                int nextCellGridSpan = nextCell.TableCellProperties.GridSpan != null ? nextCell.TableCellProperties.GridSpan.Val.Value : 0;
                if (curCellGridSpan == nextCellGridSpan &&
                curCell.InnerText == nextCell.InnerText &&
                curCell.InnerText != string.Empty &&
                nextCell.InnerText != string.Empty)
                {
                    TableCell startCell = curCell;
                    curCell.TableCellProperties.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
                    foreach (var child in nextCell.Descendants<Run>())
                    {
                        child.Parent.RemoveChild(child);
                    }
                    nextCell.TableCellProperties.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Continue };
                    nextIndex = nextIndex + 1;

                    VerticalMergeCell(tableRows, startCell, nextIndex, cIndex);
                }

            }

        }

        /// <summary>
        /// 自动纵向合并单元格
        /// </summary>
        /// <param name="tableRows">表格所有行数组</param>
        /// <param name="startCell">合并开始第一个单元格</param>
        /// <param name="nextIndex"></param>
        /// <param name="cIndex"></param>
        public void VerticalMergeCell(List<TableRow> tableRows, TableCell startCell, int nextIndex, int cIndex)
        {
            if (tableRows.Count <= nextIndex) return;

            var curRow = tableRows[nextIndex];
            var rowCells = curRow.Descendants<TableCell>().ToList();
            if (rowCells.Count <= cIndex) return;
            var nextCell = rowCells[cIndex];

            int startCellGridSpan = startCell.TableCellProperties.GridSpan != null ? startCell.TableCellProperties.GridSpan.Val.Value : 0;
            int nextCellGridSpan = nextCell.TableCellProperties.GridSpan != null ? nextCell.TableCellProperties.GridSpan.Val.Value : 0;
            if (startCellGridSpan == nextCellGridSpan &&
            startCell.InnerText == nextCell.InnerText &&
            startCell.InnerText != string.Empty &&
            nextCell.InnerText != string.Empty)
            {
                foreach (var child in nextCell.Descendants<Run>())
                {
                    child.Parent.RemoveChild(child);
                }
                nextCell.TableCellProperties.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Continue };

                nextIndex = nextIndex + 1;
                VerticalMergeCell(tableRows, startCell, nextIndex, cIndex);
            }
        }

        /// <summary>
        /// 表格自动横向合并
        /// </summary>
        /// <param name="table">表格</param>
        public void AutoHorizontalMerge(Table table)
        {
            IEnumerable<TableRow> tableRows = table.Elements<TableRow>();
            foreach (TableRow tableRow in tableRows)
            {
                HorizontalMergeCell(tableRow);
            }

        }

        /// <summary>
        /// 表格横向自动合并
        /// </summary>
        /// <param name="tableRow"></param>
        public void HorizontalMergeCell(TableRow tableRow)
        {
            List<TableCell> tableCells = tableRow.Elements<TableCell>().ToList();
            for (int i = 0; i < tableCells.Count; i++)
            {
                TableCell curCell = tableCells[i];
                TableCell nextCell = tableCells[i].NextSibling<TableCell>();

                if (nextCell != null &&
                    curCell.TableCellProperties.VerticalMerge == null &&
                    curCell.InnerText != string.Empty && nextCell.InnerText != string.Empty &&
                    curCell.InnerText == nextCell.InnerText &&
                    nextCell.TableCellProperties.VerticalMerge == null)
                {
                    int nextCellGridSpan = nextCell.TableCellProperties.GridSpan != null ? nextCell.TableCellProperties.GridSpan.Val.Value : 1;
                    int curCellGridSpan = curCell.TableCellProperties.GridSpan != null ? curCell.TableCellProperties.GridSpan.Val.Value : 1;
                    int nextCellWidth = int.Parse(nextCell.TableCellProperties.TableCellWidth.Width);
                    int curCellWidth = int.Parse(curCell.TableCellProperties.TableCellWidth.Width);

                    curCell.TableCellProperties.TableCellWidth = new TableCellWidth() { Width = (curCellWidth + nextCellWidth).ToString(), Type = TableWidthUnitValues.Dxa };
                    curCell.TableCellProperties.GridSpan = new GridSpan() { Val = curCellGridSpan + nextCellGridSpan };
                    nextCell.Remove();
                    HorizontalMergeCell(tableRow);

                }
            }

        }


        /// <summary>
        /// 合并多个word文件成一个
        /// </summary>
        /// <param name="Files">多个文件的路径列表</param>
        /// <param name="syntheticFilePath">多个文件合成一个文件存放地址，绝对路径</param>
        /// <returns></returns>
        public string MergeFiles(List<string> Files, string syntheticFilePath = null)
        {
            try
            {
                List<OpenXmlPowerTools.Source> sources = new List<OpenXmlPowerTools.Source>();

                foreach (var file in Files)
                {
                    if (!System.IO.File.Exists(file))
                    {
                        throw new Exception("文件" + file + "不存在");
                    }
                    sources.Add(new OpenXmlPowerTools.Source(new OpenXmlPowerTools.WmlDocument(file), false));
                }

                //合成文档存放绝对路径
                var mergeFileName = string.Empty;

                if (!string.IsNullOrWhiteSpace(syntheticFilePath))
                {
                    mergeFileName = syntheticFilePath;
                }
                else
                {
                    var CurrentPath = System.Environment.CurrentDirectory;//当前目录
                    var parentPath = Path.Combine(CurrentPath, "tmp");

                    var guid = System.Guid.NewGuid().ToString();
                    var fileName = string.Format("{0}.docx", guid);

                    mergeFileName = Path.Combine(parentPath, fileName);
                }

                OpenXmlPowerTools.DocumentBuilder.BuildDocument(sources, mergeFileName);

                return mergeFileName;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 检查是否为正确的文档格式
        /// </summary>
        public static void CheckIf07PlusDocx(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new Exception("无效的文件名，为空");
            }

            var ext = Path.GetExtension(fileName);
            if (!string.Equals(".docx", ext, StringComparison.OrdinalIgnoreCase))
            {
                throw new Exception("只支持docx的文档格式，当前的文档为：" + fileName);
            }
        }







    }
}