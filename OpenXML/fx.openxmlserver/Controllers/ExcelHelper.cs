using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;

namespace FX.OpenXmlServer.WordHelpers
{
    /// <summary>
    /// 定义名称实体类
    /// </summary>
    public class DefinedName
    {
        //定义名称
        public string Name;
        //定义名称所在的Sheet名
        public string SheetName;
        //引用位置
        public string Text;
        //定义名称列号
        public string Column;
        //定义名称所在行号
        public int Row;
        //多个单元格合并时终结点列好
        public string EndColumn;
        //多个单元格合并时终结点行好
        public int EndRow;
        //绑定的数据
        public Object Data;
    }


    public class ExcelHelper
    {
        //取单元格位置字母正则
        Regex charsRg = new Regex("[A-Za-z]+");

        //取单元格位置数字正则
        Regex digitRg = new Regex(@"\d+");

        // 当前打开的Excel文档对象,操作完成后,请调用Close关闭文档,释放文件占用
        public SpreadsheetDocument document = null;

        //定义名称及单元格坐标字典数据
        public List<DefinedName> DefinedNameList = new List<DefinedName>();

        //表格的起始位置
        public string from = string.Empty;

        //表格的终止位置
        public string to = string.Empty;

        //表格的列信息
        public List<string> columns = new List<string>();

        //每个sheet名及其表格数据
        public Dictionary<string, List<JObject>> TableList = new Dictionary<string, List<JObject>>();



        /// <summary>
        /// 初始化标字典数据
        /// </summary>
        /// <param name="filePath"></param>
        public void Init(string filePath, string destinationFilePath, string json)
        {
            try
            {
                File.Copy(filePath, destinationFilePath, true);
                JObject item = (JObject)JsonConvert.DeserializeObject(json);

                if (document != null)
                {
                    document.Close();
                    document.Dispose();
                }
                document = SpreadsheetDocument.Open(destinationFilePath, true);

                //获取所有自定义命名并初始化列表
                DefinedNames definedNames = document.WorkbookPart.Workbook.DefinedNames;
                //if (definedNames == null) throw new Exception("文档未定义名称");
                var WorksheetPart = document.WorkbookPart.WorksheetParts.FirstOrDefault();
                var merges = WorksheetPart.Worksheet.Descendants<MergeCell>();

                if (definedNames != null) {
                    foreach (DocumentFormat.OpenXml.Spreadsheet.DefinedName dn in definedNames)
                    {
                        if (dn.Name.Value.StartsWith("_")) continue;
                        DefinedName definedName = new DefinedName();
                        definedName.Name = dn.Name.Value;
                        int index = dn.Text.LastIndexOf("!");
                        definedName.SheetName = dn.Text.Substring(0, index).Trim('\'');
                        definedName.Text = dn.Text;
                        string Raw = dn.Text.Substring(index + 1, dn.Text.Length - index - 1);
                        definedName.Column = definedName.EndColumn = Raw.Split('$')[1];
                        definedName.Row = definedName.EndRow = int.Parse(Raw.Split('$')[2]);
                        //if (!item.ContainsKey(definedName.Name)) throw new Exception(string.Format("JSON数据不存在【{0}】字段", definedName.Name));
                        definedName.Data = item.ContainsKey(definedName.Name) ? item[definedName.Name] : string.Empty;

                        foreach (var merge in merges)
                        {
                            if (merge.Reference.Value.StartsWith(definedName.Column + definedName.Row))
                            {
                                string endCell = merge.Reference.Value.Split(':')[1];

                                definedName.EndColumn = charsRg.Match(endCell).Value;
                                definedName.EndRow = int.Parse(digitRg.Match(endCell).Value);
                            }
                        }

                        if (definedName.Name == "PageInfo")
                        {
                            definedName.Data = "第1页 共1页";
                        }

                        DefinedNameList.Add(definedName);
                    }
                }



                //遍历文档第一个sheet首行，查找用户定义的配置信息
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>();
                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                Row row = sheetData.Elements<Row>().FirstOrDefault();

                foreach (Cell cell in row)
                {
                    string value = GetCellValue(cell);
                    if (value.IndexOf("table") > -1)
                    {
                        try
                        {
                            JObject config = (JObject)JsonConvert.DeserializeObject(value);
                            from = (string)config["table"]["SampleList"]["from"];
                            to = (string)config["table"]["SampleList"]["to"];
                            JArray array = (JArray)config["table"]["Columns"];
                            if (array != null)
                            {
                                for (int i = 0; i < array.Count(); i++)
                                {
                                    columns.Add((string)array[i]);
                                }
                            }

                            break;
                        }
                        catch (Exception e)
                        {
                            throw new Exception("配置异常！");
                        }

                    }

                }

                string sourceSheetName = document.WorkbookPart.Workbook.Descendants<Sheet>().First().Name;

                //表格列配置存在且数据包含SampleList
                if (columns.Count() != 0 && item.ContainsKey("SampleList"))
                {
                    JArray sampleList = (JArray)item["SampleList"];
                    int sum = int.Parse(digitRg.Match(to).Value) - int.Parse(digitRg.Match(from).Value) + 1;
                    float length = (float)sampleList.Count();

                    //处理第一个sheet的表格
                    List<JObject> ls = new List<JObject>();
                    for (int i = 0; i < sum; i++)
                    {
                        if (sampleList.Count() == 0) break;
                        ls.Add((JObject)sampleList[0]);
                        sampleList.RemoveAt(0);
                    }
                    TableList.Add(sourceSheetName, ls);

                    
                    int x = (int)(length / sum);
                    if (x == 0) return;
                    float y = (length / sum) - (float)x;
                    if (y > 0f) x++;




                    //分页之后复制定义名称及每一个sheet的表格
                    List<DefinedName> temp = new List<DefinedName>();
                    for (int i = 0; i < x - 1; i++)
                    {
                        string sheetName = CopyWorksheet(i);
                        //复制一遍自定义名称
                        var list = DefinedNameList.Where(e => e.SheetName == sourceSheetName);
                        foreach (var dn in list)
                        {
                            DefinedName ndn = new DefinedName();
                            ndn.Name = dn.Name + "_" + i;
                            ndn.Column = dn.Column;
                            ndn.Row = dn.Row;
                            ndn.EndColumn = dn.EndColumn;
                            ndn.EndRow = dn.EndRow;
                            ndn.SheetName = sheetName;
                            ndn.Text = dn.Text.Replace(sourceSheetName, sheetName);
                            ndn.Data = dn.Data;
                            if (dn.Name == "PageInfo") {
                                dn.Data = "第1页 共" + x + "页";
                                ndn.Data = "第" + (i + 2) + "页 共" + x + "页";
                            }



                            temp.Add(ndn);
                        }

                        //复制一遍表格
                        List<JObject> lst = new List<JObject>();
                        for (int j = 0; j < sum; j++)
                        {
                            if (sampleList.Count() == 0) break;
                            lst.Add((JObject)sampleList[0]);
                            sampleList.RemoveAt(0);
                        }
                        TableList.Add(sheetName, lst);
                    }
                    DefinedNameList = DefinedNameList.Concat(temp).ToList();

                }












                //根据JSON数据计算是否分页
                /*                foreach (var ele in item)
                                {
                                    var Key = ele.Key;
                                    var Val = ele.Value;
                                    if (Val is JArray)
                                    {

                                        int sum = int.Parse(digitRg.Match(Key.Split("_")[1]).Value) - int.Parse(digitRg.Match(Key.Split("_")[0]).Value) + 1;
                                        float length = (float)Val.Count();
                                        int x = (int)(length/sum);
                                        if (x == 0) return;
                                        float y = (length / sum) - (float)x;
                                        if (y > 0f) x++;

                                        JArray a = new JArray();
                                        for (int i = 0; i < sum; i++) {
                                            a.Add(Val[0]);
                                            ((JArray)Val).RemoveAt(0);
                                        }

                                        var dflist = DefinedNameList.Where(e => e.SheetName == sourceSheetName && e.Name == Key);
                                        dflist.First().Data = a;

                                        List<DefinedName> temp = new List<DefinedName>();
                                        for (int i = 0; i < x - 1; i ++) {
                                            var list = DefinedNameList.Where(e => e.SheetName == sourceSheetName);

                                            JArray ja = new JArray();
                                            for (int j = 0; j < sum; j++)
                                            {
                                                if (((JArray)Val).Count() == 0) break;
                                                ja.Add(Val[0]);
                                                ((JArray)Val).RemoveAt(0);
                                            }


                                            string sheetName = CopyWorksheet(i);
                                            foreach (var dn in list)
                                            {
                                                DefinedName ndn = new DefinedName();
                                                ndn.Name = dn.Name + "_" + i;
                                                ndn.Column = dn.Column;
                                                ndn.Row = dn.Row;
                                                ndn.EndColumn = dn.EndColumn;
                                                ndn.EndRow = dn.EndRow;
                                                ndn.SheetName = sheetName;
                                                ndn.Text = dn.Text.Replace(sourceSheetName, sheetName);
                                                if (dn.Name == Key)
                                                {
                                                    ndn.Data = ja;
                                                }
                                                else
                                                {
                                                    ndn.Data = dn.Data;
                                                }
                                                temp.Add(ndn);
                                            }
                                        }

                                        DefinedNameList = DefinedNameList.Concat(temp).ToList();
                                    }
                                }*/
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 处理过程
        /// </summary>
        /// <param name="json"></param>
        public void ExcelProcess()
        {
            try
            {
                foreach (DefinedName definedName in DefinedNameList)
                {
                    /*                    if (definedName.Data is JArray)
                                        {
                                            int r = definedName.Row;
                                            foreach (var array in (JArray)definedName.Data) {
                                                string c = definedName.Column;
                                                foreach (var val in (JArray)array)
                                                {
                                                    c = SetArray(definedName.SheetName, c, r, val.ToString());
                                                }
                                                ++ r;
                                            }
                                        }
                                        else {*/
                    if (File.Exists(definedName.Data.ToString()))
                    {
                        InsertImage(definedName.SheetName, definedName.Data.ToString(), definedName.Row - 1, ToIndex(definedName.Column) - 1, definedName.EndRow - 1, ToIndex(definedName.EndColumn) - 1);
                    }
                    else
                    {
                        InsertText(definedName.SheetName, definedName.Data.ToString(), definedName.Column, (uint)definedName.Row);
                    }
                    //}

                }

                foreach (KeyValuePair<string, List<JObject>> kv in TableList)
                {
                    string sheetName = kv.Key;
                    List<JObject> list = kv.Value;
                    IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName);
                    WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                    int r = int.Parse(digitRg.Match(from).Value);
                    for (int i = 0; i < list.Count(); i++)
                    {
                        JObject ele = list[i];
                        /*                        List<string> ls = new List<string>();
                                                for (int j = 0; j < columns.Count; j ++) {
                                                    if (columns[j] == string.Empty) {
                                                        ls.Add(string.Empty);
                                                        continue;
                                                    }
                                                    ls.Add((string)ele[columns[j]]);
                                                }*/

                        //string c = charsRg.Match(from).Value;

                        int Startc = ToIndex(charsRg.Match(from).Value);
                        int Endc = ToIndex(charsRg.Match(to).Value);
                        Row row = sheetData.Elements<Row>().Where(e => e.RowIndex.Value == r).FirstOrDefault();

                        //收集所有需要处理的cell
                        List<string> cells = new List<string>();
                        IEnumerable<Cell> cs = row.Elements<Cell>().Where(cell => ToIndex(charsRg.Match(cell.CellReference.Value).Value) >= Startc
                                                                    && ToIndex(charsRg.Match(cell.CellReference.Value).Value) <= Endc);
                        foreach (Cell cell in cs)
                        {
                            string reference = cell.CellReference.Value;
                            int cellRow = int.Parse(digitRg.Match(reference).Value);
                            string cellColumn = charsRg.Match(reference).Value;
                            cells.Add(GetTopLeftCell(cellColumn, cellRow));
                        }
                        cells = cells.Distinct().ToList();

                        //填充一行
                        for (int j = 0; j < cells.Count(); j++)
                        {
                            uint cellRow = uint.Parse(digitRg.Match(cells[j]).Value);
                            string cellColumn = charsRg.Match(cells[j]).Value;
                            if (columns[j] == string.Empty)
                            {
                                InsertText(sheetName, "", cellColumn, cellRow);
                                continue;
                            }
                            InsertText(sheetName, (string)ele[columns[j]], cellColumn, cellRow);
                        }

                        ++r;
                    }
                }
                //document.save()
            }
            catch (Exception e)
            {
                throw e;
            }

        }




        public string GetTopLeftCell(string Column, int Row)
        {
            int ColumnId = ToIndex(Column);
            var WorksheetPart = document.WorkbookPart.WorksheetParts.FirstOrDefault();
            var merges = WorksheetPart.Worksheet.Descendants<MergeCell>();

            foreach (var merge in merges)
            {
                string startCell = merge.Reference.Value.Split(':')[0];
                string endCell = merge.Reference.Value.Split(':')[1];
                if ((Column + Row) == startCell) return startCell;
                int startCellColumne = ToIndex(charsRg.Match(startCell).Value);
                int startCellRow = int.Parse(digitRg.Match(startCell).Value);
                int endCellColumne = ToIndex(charsRg.Match(endCell).Value);
                int endCellRow = int.Parse(digitRg.Match(endCell).Value);
                if ((ColumnId >= startCellColumne && ColumnId <= endCellColumne) &&
                    (Row >= startCellRow && Row <= endCellRow))
                    return startCell;
            }
            return Column + Row;
        }



        /// <summary>
        /// 单元格合并时，判断哪些单元格不需要插入数据
        /// </summary>
        /// <param name="Column"></param>
        /// <param name="Row"></param>
        /// <returns></returns>
        public bool BetweenTwoCell(string Column, int Row)
        {
            int ColumnId = ToIndex(Column);
            var WorksheetPart = document.WorkbookPart.WorksheetParts.FirstOrDefault();
            var merges = WorksheetPart.Worksheet.Descendants<MergeCell>();
            foreach (var merge in merges)
            {
                string startCell = merge.Reference.Value.Split(':')[0];
                string endCell = merge.Reference.Value.Split(':')[1];
                if ((Column + Row) == startCell) return false;
                int startCellColumne = ToIndex(charsRg.Match(startCell).Value);
                int startCellRow = int.Parse(digitRg.Match(startCell).Value);
                int endCellColumne = ToIndex(charsRg.Match(endCell).Value);
                int endCellRow = int.Parse(digitRg.Match(endCell).Value);
                if ((ColumnId >= startCellColumne && ColumnId <= endCellColumne) &&
                    (Row >= startCellRow && Row <= endCellRow))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// 多行数据插入
        /// </summary>
        /// <param name="Column"></param>
        /// <param name="Row"></param>
        /// <param name="val"></param>
        /// <returns></returns>
        public void SetArray(string worksheetName, string Column, int Row, List<string> ls)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
            if (sheets.Count() == 0)
            {
                throw new Exception("不存在指定的Sheet");
            }
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            int ColumnIndex = ToIndex(Column);
            foreach (var val in ls) {
                var newColumn = ToName(ColumnIndex);
                var cell = InsertCellInWorksheet(newColumn, (uint)Row, worksheetPart);
                int index = InsertSharedStringItem(val);
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                ColumnIndex++;
            }





/*            if (ls.Count == 0) return;
            if (BetweenTwoCell(Column, Row))
            {
                int c = ToIndex(Column);
                string newCol = ToName(++c);
                SetArray(worksheetName, newCol, Row, ls);
                return;
            }
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
            if (sheets.Count() == 0)
            {
                throw new Exception("不存在指定的Sheet");
            }
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            //var sheetPart = document.WorkbookPart.WorksheetParts.FirstOrDefault();
            var cell = InsertCellInWorksheet(Column, (uint)Row, worksheetPart);//worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference.Value == Column + Row).First();
            if (ls.Count == 0) return;
            string val = ls[0];
            ls.RemoveAt(0);
            int index = InsertSharedStringItem(val);
            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            int ColumnId = ToIndex(Column);
            string cc = ToName(++ColumnId);
            SetArray(worksheetName, cc, Row, ls);
            if (ls.Count == 0) return;*/
        }



        public string CopyWorksheet(int index)
        {
            string sourceSheetName = document.WorkbookPart.Workbook.Descendants<Sheet>().First().Name;
            var sourceSheetPart = document.WorkbookPart.WorksheetParts.FirstOrDefault(); //GetWorksheetPart(sourceSheetName);

            var tempSheet = SpreadsheetDocument.Create(new MemoryStream(), document.DocumentType);
            var tempWorkbookPart = tempSheet.AddWorkbookPart();

            var tempWorksheetPart = tempWorkbookPart.AddPart<WorksheetPart>(sourceSheetPart);

            //Add cloned sheet and all associated parts to workbook

            var clonedSheetPart = document.WorkbookPart.AddPart<WorksheetPart>(tempWorksheetPart);

            //Table definition parts are somewhat special and need unique ids…so let’s make an id based on count
            var numTableDefParts = sourceSheetPart.GetPartsCountOfType<TableDefinitionPart>();
            var tableId = numTableDefParts;
            //Clean up table definition parts (tables need unique ids)
            if (numTableDefParts != 0)
            {
                //Every table needs a unique id and name
                foreach (TableDefinitionPart tableDefPart in clonedSheetPart.TableDefinitionParts)
                {
                    tableId++;

                    tableDefPart.Table.Id = (uint)tableId;
                    tableDefPart.Table.DisplayName = "CopiedTable" + tableId;
                    tableDefPart.Table.Name = "CopiedTable" + tableId;
                    tableDefPart.Table.Save();
                }
            }

            //There should only be one sheet that has focus
            var views = clonedSheetPart.Worksheet.GetFirstChild<SheetViews>();
            if (views != null)
            {
                views.Remove();
                clonedSheetPart.Worksheet.Save();
            }

            // last step is to add a reference to the added worksheet in the main workbook part
            var sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            var clonedAsSheetName = sourceSheetName + (index + 1);// (sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1);
            //ValidSheetName(sheets, ref clonedAsSheetName);

            var copiedSheet = new Sheet
            {
                Name = clonedAsSheetName,
                Id = document.WorkbookPart.GetIdOfPart(clonedSheetPart),
                SheetId = (uint)sheets.ChildElements.Count + 1
            };

            sheets.AppendChild(copiedSheet);

            //copy sheet added DefinedName
            var definedNameList = DefinedNameList.Where(s => s.SheetName == sourceSheetName).ToList();
            foreach (var definedName in definedNameList)
            {
                DocumentFormat.OpenXml.Spreadsheet.DefinedName dn = new DocumentFormat.OpenXml.Spreadsheet.DefinedName()
                {
                    Name = definedName.Name + "_" + index,
                    Text = string.Format("{0}!${1}${2}", clonedAsSheetName, definedName.Column, definedName.Row),
                };


                document.WorkbookPart.Workbook.DefinedNames.Append(dn);
            }

            //Save Changes
            document.WorkbookPart.Workbook.Save();

            return clonedAsSheetName;

        }


        /// <summary>
        /// 关闭文档释放资源
        /// </summary>
        public void Close()
        {
            try
            {
                document.WorkbookPart.Workbook.Save();
                document.Close();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 获取单元格的值
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public string GetCellValue(Cell cell)
        {
            string value = null;
            if (cell != null)
            {
                value = cell.InnerText;
                if (cell.DataType != null)
                {
                    switch (cell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            var stringTable = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            if (stringTable != null)
                            {
                                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }
            return value;
        }


        /// <summary>
        /// 往单元格插入数据
        /// </summary>
        /// <param name="text"></param>
        /// <param name="columnName"></param>
        /// <param name="rowIndex"></param>
        public void InsertText(string worksheetName, string text, string columnName, uint rowIndex)
        {

            int index = InsertSharedStringItem(text);
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName); //document.WorkbookPart.Workbook.Descendants<Sheet>();
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            Cell cell = InsertCellInWorksheet(columnName, rowIndex, worksheetPart);
            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            worksheetPart.Worksheet.Save();
        }


        /// <summary>
        /// 插入共享字符串，并获取该字符串唯一索引
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private int InsertSharedStringItem(string text)
        {
            SharedStringTablePart shareStringPart;
            if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }
            int i = 0;
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();
            return i;
        }

        /// <summary>
        /// 在Worksheet插入单元格
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="rowIndex"></param>
        /// <param name="worksheetPart"></param>
        /// <returns></returns>
        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                /*                Cell refCell = null;
                                foreach (Cell cell in row.Elements<Cell>())
                                {
                                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                                    {
                                        refCell = cell;
                                        break;
                                    }
                                }
                                Cell newCell = new Cell() { CellReference = cellReference };
                                row.InsertBefore(newCell, refCell);*/
                Cell newCell = new Cell() { CellReference = cellReference };
                int index = ToIndex(columnName) - 1;
                row.InsertAt(newCell, index);
                worksheet.Save();
                return newCell;
            }
        }

        /// <summary>
        /// 获取行号的数字形式
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static int ToIndex(string columnName)
        {
            if (!Regex.IsMatch(columnName.ToUpper(), @"[A-Z]+")) { throw new Exception("invalid parameter"); }

            int index = 0;
            char[] chars = columnName.ToUpper().ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                index += ((int)chars[i] - (int)'A' + 1) * (int)Math.Pow(26, chars.Length - i - 1);
            }
            return index;
        }

        /// <summary>
        /// 根据行号数字形式获取字母索引
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public static string ToName(int index)
        {
            if (--index < 0) { throw new Exception("invalid parameter"); }

            List<string> chars = new List<string>();
            do
            {
                if (chars.Count > 0) index--;
                chars.Insert(0, ((char)(index % 26 + (int)'A')).ToString());
                index = (int)((index - index % 26) / 26);
            } while (index > 0);

            return String.Join(string.Empty, chars.ToArray());
        }


        /// <summary>
        /// 往sheet表格渲染图片
        /// </summary>
        /// <param name="imgPath"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="endRowIndex"></param>
        /// <param name="endColumnIndex"></param>
        public void InsertImage(string worksheetName, string imgPath, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {


            ImagePartType ipt;
            switch (Path.GetExtension(imgPath).TrimStart('.').ToLower())
            {
                case "png":
                    ipt = ImagePartType.Png;
                    break;
                case "jpg":
                case "jpeg":
                    ipt = ImagePartType.Jpeg;
                    break;
                case "gif":
                    ipt = ImagePartType.Gif;
                    break;
                default:
                    return;
            }

            //var wsPart = document.WorkbookPart.WorksheetParts.FirstOrDefault();
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);

            WorksheetPart wsPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            int drawingPartId = GetNextRelationShipID(wsPart);

            Drawing drawing1 = new Drawing() { Id = "rId" + drawingPartId.ToString() };
            if (wsPart.VmlDrawingParts == null)
            {
                wsPart.Worksheet.Append(drawing1);
            }
            else
            {
                var ds = wsPart.Worksheet.Descendants<Drawing>();

                if (!ds.Any())
                {
                    wsPart.Worksheet.Append(drawing1);
                }
            }


            DrawingsPart drawingsPart;
            ImagePart imagePart;

            if (wsPart.DrawingsPart == null)
            {
                drawingsPart = wsPart.AddNewPart<DrawingsPart>("rId" + drawingPartId.ToString());

                var imgId = "rId" + (drawingsPart.GetPartsOfType<ImagePart>().Count() + 1);

                GenerateDrawingsPart1Content(drawingsPart, imgId, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);

                imagePart = drawingsPart.AddImagePart(ipt, imgId);
            }
            else
            {
                drawingsPart = wsPart.DrawingsPart;

                var imgId = "rId" + (drawingsPart.GetPartsOfType<ImagePart>().Count() + 1);

                imagePart = drawingsPart.AddImagePart(ipt, imgId);

                GenerateDrawingsPart1Content(drawingsPart, imgId, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);

            }

            using (var imgFs = new FileStream(imgPath, FileMode.Open))
            {
                imagePart.FeedData(imgFs);
            }

        }


        private void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1, string imgId, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            var worksheetDrawing1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var twoCellAnchor1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor() { EditAs = DocumentFormat.OpenXml.Drawing.Spreadsheet.EditAsValues.OneCell };

            var fromMarker1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker();
            var columnId1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId();
            columnId1.Text = startColumnIndex.ToString();
            var columnOffset1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset();
            columnOffset1.Text = "38100";
            var rowId1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId();
            rowId1.Text = startRowIndex.ToString();
            var rowOffset1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset();
            rowOffset1.Text = "0";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            var toMarker1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker();
            var columnId2 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId();
            columnId2.Text = endColumnIndex.ToString();
            var columnOffset2 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset();
            columnOffset2.Text = "542925";
            var rowId2 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId();
            rowId2.Text = endRowIndex.ToString();
            var rowOffset2 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset();
            rowOffset2.Text = "161925";

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            var picture1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture();

            var nonVisualPictureProperties1 = new NonVisualPictureProperties();
            var nonVisualDrawingProperties1 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" };

            var nonVisualPictureDrawingProperties1 = new NonVisualPictureDrawingProperties();
            var pictureLocks1 = new DocumentFormat.OpenXml.Drawing.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            var blipFill1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill();

            var blip1 = new DocumentFormat.OpenXml.Drawing.Blip() { Embed = imgId, CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print };
            blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            var blipExtensionList1 = new DocumentFormat.OpenXml.Drawing.BlipExtensionList();

            var blipExtension1 = new DocumentFormat.OpenXml.Drawing.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            var useLocalDpi1 = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            var stretch1 = new DocumentFormat.OpenXml.Drawing.Stretch();
            var fillRectangle1 = new DocumentFormat.OpenXml.Drawing.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            var shapeProperties1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties();

            var transform2D1 = new DocumentFormat.OpenXml.Drawing.Transform2D();
            var offset1 = new DocumentFormat.OpenXml.Drawing.Offset() { X = 1257300L, Y = 762000L };
            var extents1 = new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 2943225L, Cy = 2257425L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            var presetGeometry1 = new DocumentFormat.OpenXml.Drawing.PresetGeometry() { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle };
            var adjustValueList1 = new DocumentFormat.OpenXml.Drawing.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            var clientData1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(picture1);
            twoCellAnchor1.Append(clientData1);

            if (drawingsPart1.WorksheetDrawing == null)
            {
                worksheetDrawing1.Append(twoCellAnchor1);
                drawingsPart1.WorksheetDrawing = worksheetDrawing1;
            }
            else
            {
                drawingsPart1.WorksheetDrawing.Append(twoCellAnchor1);
            }
        }

        private int GetNextRelationShipID(WorksheetPart sheet1)
        {
            int nextId = 0;
            List<int> ids = new List<int>();
            foreach (IdPartPair part in sheet1.Parts)
            {
                ids.Add(int.Parse(part.RelationshipId.Replace("rId", string.Empty)));
            }
            if (ids.Count > 0)
            {
                nextId = ids.Max() + 1;
            }
            else
            {
                nextId = 1;
            }

            return nextId;
        }



    }
}
