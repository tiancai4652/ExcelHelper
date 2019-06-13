using Aspose.Cells;
using Aspose.Cells.Properties;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Shaco_Aspose
{
    public class AsposeExcelApplication
    {
        #region 属性

        public Workbook Workbook { get; set; }

        #endregion

        #region 构造函数

       
        public AsposeExcelApplication()
        {
            Workbook = new Workbook();
        }

        /// <summary>
        /// 打开指定的文件
        ///  Aspose.Cells will automatically detect the file format type.
        /// </summary>
        /// <param name="fileName"></param>
        public AsposeExcelApplication(string fileName)
        {
            Workbook = new Workbook(fileName);
            Workbook.Settings.CreateCalcChain = false;
        }


        #endregion

      

        #region 工作薄操作

        /// <summary>
        /// Excel工作薄设置文档属性
        /// </summary>
        /// <param name="key"></param>
        /// Subject
        /// Author
        /// Keywords
        /// Comments
        /// Template
        /// Last Author
        /// Revision Number
        /// Application Name
        /// Last Print Date
        /// Creation Date
        /// Last Save Time
        /// Total Editing Time
        /// Number of Pages
        /// Number of Words
        /// Number of Characters
        /// Security
        /// Category
        /// Format
        /// Manager
        /// Company
        /// Number of Bytes
        /// Number of Lines
        /// Number of Paragraphs
        /// Number of Slides
        /// Number of Notes
        /// Number of Hidden Slides
        /// Number of Multimedia Clips
        /// <param name="value"></param>
        public void SetWorkbookDocementProperties(string key, object value)
        {
            var collection = Workbook.Worksheets.BuiltInDocumentProperties;
            if (collection.Contains(key))
            {
                DocumentProperty doc = collection[key];
                doc.Value = value;
            }
        }

        /// <summary>
        /// 获取Excel工作簿文档属性
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public object GetWorkbookDocementProperties(string key)
        {
            var collection = Workbook.Worksheets.BuiltInDocumentProperties;
            if (collection.Contains(key))
            {
                DocumentProperty doc = collection[key];
                switch(doc.Type)
                {
                    case PropertyType.Boolean:
                        return doc.ToBool();
                    case PropertyType.DateTime:
                        return doc.ToDateTime();
                    case PropertyType.Double:
                        return doc.ToDouble();
                    case PropertyType.Number:
                        return doc.ToInt();
                    case PropertyType.String:
                        return doc.ToString();
                    default:
                        return null;

                }
            }
            return null;
        }

        /// <summary>
        /// Excel工作簿添加自定义属性
        /// </summary>
        /// <param name="propName">属性名称</param>
        /// <param name="propValue">属性值</param>
        public void SetWorkbookDocCustomProperties(string propName, string propValue)
        {
            Workbook.Worksheets.CustomDocumentProperties.Add(propName, propValue);
        }

        /// <summary>
        /// 获取Excel工作簿自定义属性
        /// </summary>
        /// <param name="propName"></param>
        /// <returns></returns>
        public object GetWorkbookDocCustomProperties(string propName)
        {
            var collection = Workbook.Worksheets.CustomDocumentProperties;
            if (collection.Contains(propName))
            {
                DocumentProperty doc = collection[propName];
                return doc.ToString();
            }
            return null;
        }

        /// <summary>
        /// 保存
        /// </summary>
        public void Save()
        {
            Workbook.Save(Workbook.FileName);
        }

        /// <summary>
        /// 保存
        /// </summary>
        /// <param name="path"></param>
        public void Save(string path)
        {
            Workbook.Save(path);
        }
        #endregion

        #region Sheet

        /// <summary>
        /// 创建工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        public void CreateSheet(string sheetName)
        {
            if (!(Workbook.Worksheets.Find(t => t.Name.Equals(sheetName))!=null))
            {
                Workbook.Worksheets.Add(sheetName);
            }
        }

        /// <summary>
        /// 判断工作表是否存在
        /// </summary>
        /// <param name="sheetName">工作表名</param>
        /// <returns></returns>
        public bool IsExistSheet(string sheetName)
        {
            return Workbook.Worksheets.Exists(t=>t.Name.Equals(sheetName));
        }

        /// <summary>
        /// 获取所有未隐藏的sheet
        /// </summary>
        /// <param name="postfix"></param>
        public IEnumerable<Worksheet> GetDisplaySheets()
        {
            return Workbook.Worksheets.Where(t => t.IsVisible == true);
        }

        /// <summary>
        /// 复制当前工作簿的Sheet到当前工作簿
        /// </summary>
        /// <param name="sourceSheetName"></param>
        /// <param name="targetSheetName"></param>
        public void CopySheetInSelf(string sourceSheetName,string targetSheetName)
        {
           int index= Workbook.Worksheets.AddCopy(sourceSheetName);
            Workbook.Worksheets[index].Name = targetSheetName;
        }

        /// <summary>
        /// 复制其他工作簿的Sheet到当前工作簿
        /// </summary>
        public void CopySheetToAnotherWorkbook(Worksheet sourceWorkSheet,string targetSheetName)
        {
            CreateSheet(targetSheetName);
            var sheet = Workbook.Worksheets[targetSheetName];
            sheet.Copy(sourceWorkSheet);
        }

        /// <summary>
        /// 重命名一张工作表
        /// </summary>
        /// <param name="oldSheetName"></param>
        /// <param name="newSheetName"></param>
        public void RenameSheet(string oldSheetName, string newSheetName)
        {
            var sheet = Workbook.Worksheets.Find(t => t.Name.Equals(oldSheetName));
            if (sheet != null)
            {
                sheet.Name = newSheetName;
            }
        }

        /// <summary>
        /// 移除指定的工作表
        /// </summary>
        /// <param name="sheetName"></param>
        public void RemoveSheet(string sheetName)
        {
            var sheet = Workbook.Worksheets.Find(t => t.Name.Equals(sheetName));
            if (sheet != null)
            {
                Workbook.Worksheets.RemoveAt(sheetName);
            }
        }

        /// <summary>
        /// 缩放工作表
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="numerator">分子</param>
        /// <param name="denominator">分母</param>
        public void ZoomSheet(string sheetName, int numerator, int denominator)
        {
            if (denominator == 0)
            {
                return;
            }
            var sheet = Workbook.Worksheets.Find(t => t.Name.Equals(sheetName));
            if (sheet != null)
            {
                sheet.Zoom = numerator*100/ denominator;
            }
        }

        /// <summary>
        /// 获取隐藏的数据表集合
        /// </summary>
        /// <returns></returns>
        public List<string> GetHideSheets()
        {
            return Workbook.Worksheets.Where(t => t.IsVisible == false).Select(t => t.Name).ToList();
        }

        /// <summary>
        /// 设置Sheet显示还是隐藏
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="IsVisible"></param>
        public void SetSheetVisible(string sheetName,bool IsVisible)
        {
            var sheet = Workbook.Worksheets.Find(t => t.Name.Equals(sheetName));
            if (sheet != null)
            {
                sheet.IsVisible = IsVisible;
            }
        }

        /// <summary>
        /// 设置已存在的工作表位置
        /// </summary>
        /// <param name="sheetName">工作表表名</param>
        /// <param name="index">要设置的位置，从0开始</param>
        /// <returns>是否成功，工作表不存在时返回false</returns>
        public bool SetSheetIndex(string sheetName, int index)
        {
            var sheet = Workbook.Worksheets[sheetName];
            if (sheet == null || Workbook.Worksheets.Count < index + 1)
            {
                return false;
            }
            else
            {
                sheet.MoveTo(index);
                return true;

            }
        }

        /// <summary>
        /// 设置活动工作表
        /// </summary>
        /// <param name="SheetName"></param>
        public void SetActiveSheet(string sheetName)
        {
            var sheet = Workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                return ;
            }
            Workbook.Worksheets.ActiveSheetIndex = sheet.Index;
        }

        /// <summary>
        /// 设置活动工作表
        /// </summary>
        public void SetActiveSheet(int index)
        {
            if (Workbook.Worksheets.Count >= index + 1)
            {
                Workbook.Worksheets.ActiveSheetIndex = index;
            }
        }

        /// <summary>
        /// 计算工作表中的公式
        /// </summary>
        /// <param name="sheetName"></param>
        public void Recalculation(string sheetName)
        {
            var sheet = Workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                return;
            }
            //sheet.CalculateFormula(new CalculationOptions(),true);
            sheet.CalculateFormula(true, true, null);
        }

        /// <summary>
        /// 从工作表里获取数据
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="endRowIndex">如果为0，执行到最后一行</param>
        /// <param name="endColumnIndex"></param>
        /// <returns></returns>
        public DataTable GetSheetData(string sheetName, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            if (endRowIndex != 0 && startRowIndex > endRowIndex && startColumnIndex > endColumnIndex)
                return null;
            var sheet = Workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                return null;
            }
            DataTable dt = new DataTable();
            for (int j = startColumnIndex; j <= endColumnIndex; j++)
            {
                dt.Columns.Add("Column" + j.ToString());
            }
            for (int i = 0; i < sheet.Cells.Rows.Count; i++)
            {
                if (sheet.Cells.Rows[i].Index < startRowIndex || sheet.Cells.Rows[i].Index > endRowIndex)
                {
                    continue;
                }
                DataRow dr = dt.NewRow();
                for (int c = startColumnIndex; c <= endColumnIndex; c++)
                {
                    var cell = sheet.Cells.Rows[i][c];

                    if (cell == null)
                    {
                        dr[c - startColumnIndex] = string.Empty;
                    }
                    else
                    {
                        switch (cell.Type)
                        {
                            case CellValueType.IsBool:
                                dr[c - startColumnIndex] = cell.StringValue;
                                break;
                            case CellValueType.IsDateTime:
                                dr[c - startColumnIndex] = cell.DateTimeValue;
                                break;
                            case CellValueType.IsNumeric:
                                dr[c - startColumnIndex] = cell.FloatValue;
                                break;
                            case CellValueType.IsString:
                                dr[c - startColumnIndex] = cell.StringValue;
                                break;
                            default:
                                dr[c - startColumnIndex] = string.Empty;
                                break;
                        }
                       
                    }
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        /// <summary>
        /// 保护工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="password">密码</param>
        public void ProtectSheet(string sheetName, string password, string oldPassword = null)
        {
            var sheet = Workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                return;
            }
            sheet.Protect(ProtectionType.All, password, oldPassword);
        }

        
        #endregion

        #region Row/Column

        /// <summary>
        /// 获取工作表最后行号
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public int GetSheetLastRowNumber(string sheetName)
        {
            var sheet = Workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                return 0;
            }
            return sheet.Cells.MaxDataRow;
        }

        /// <summary>
        /// 获取工作表某一行中最后的单元格列号（From 0）
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public int GetRowLastCellNumber(string sheetName, int rowIndex)
        {
            var sheet = Workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                return 0;
            }
            return sheet.Cells.Rows[rowIndex].LastCell.Column;
        }

        /// <summary>
        /// 插入空列
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="cellInsertIndex"></param>
        public void InsertEmptyColumn(string sheetName, int emptyColumnIndex)
        {
            var sheet = Workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                return;
            }
            sheet.Cells.InsertColumn(emptyColumnIndex,true);
        }

        // <summary>
        /// 插入空行
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="insertRowIndex">待插入行的索引位置</param>
        public void InsertEmptyRow(string sheetName, int insertRowIndex)
        {
            var sheet = Workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                return;
            }
            sheet.Cells.InsertRow(insertRowIndex);
        }

        /// <summary>
        /// 设置行高(in unit of Points.)
        /// </summary>
        public void SetRowHeight(string sheetName,int rowIndex,double height)
        {
            var sheet = Workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                return;
            }
            if (sheet.Cells.Rows.Count >= rowIndex + 1)
            {
                sheet.Cells.Rows[rowIndex].Height = height;
            }
        }


        #endregion

        #region 单元格

        /// <summary>
        /// 向单元格内填充值
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="rowIndex"></param>
        /// <param name="ColumnIndex"></param>
        /// <param name="value"></param>
        public void SetCellValue(string sheetName, int rowIndex, int ColumnIndex, object value)
        {
            var cell = Workbook.Worksheets[sheetName].Cells[rowIndex, ColumnIndex];
            if (value == null || (value is string && string.IsNullOrEmpty(value as string)))
                return;
            switch (value.GetType().ToString())
            {
                case "System.String":
                    cell.PutValue(value.ToString());
                    break;
                case "System.DateTime":
                    cell.PutValue(DateTime.Parse(value.ToString()).ToShortDateString());
                    break;
                case "System.Boolean":
                    bool boolV = false;
                    bool.TryParse(value.ToString(), out boolV);
                    cell.PutValue(boolV);
                    break;
                case "System.Int16":
                case "System.Int32":
                case "System.Int64":
                case "System.Byte":
                    int intV = 0;
                    int.TryParse(value.ToString(), out intV);
                    cell.PutValue(intV);
                    break;
                case "System.Decimal":
                case "System.Double":
                case "System.Single":
                    if ((double)value == double.NaN)
                        cell.PutValue("");
                    else
                    {
                        double doubV = 0;
                        double.TryParse(value.ToString(), out doubV);
                        cell.PutValue(doubV);
                        //if (p.RoundDigits != int.MinValue)
                        //{
                        //    SetNumericCell(cell, p.RoundDigits);
                        //}
                    }
                    break;
                case "System.DBNull":
                    cell.PutValue("");
                    break;
                default:
                    cell.PutValue(value.ToString());
                    break;
            }

        }

        /// <summary>
        /// 获取单元格值
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="rowIndex"></param>
        /// <param name="ColumnIndex"></param>
        /// <returns></returns>
        public object GetCellValue(string sheetName, int rowIndex, int ColumnIndex)
        {
            var cell = Workbook.Worksheets[sheetName].Cells[rowIndex, ColumnIndex];
            switch (cell.Type)
            {
                case CellValueType.IsBool:
                    return cell.BoolValue;
                case CellValueType.IsDateTime:
                    return cell.DateTimeValue;
                case CellValueType.IsNumeric:
                    return cell.FloatValue;
                case CellValueType.IsString:
                    return cell.StringValue;
                default:
                    return null;
            }
        }

        /// <summary>
        /// 给单元格设置公式
        /// 计算会消耗性能，可以添加所有公示后再计算( Workbook.CalculateFormula();)
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="cellFormulaString"></param>
        public void SetCellFormula(string sheetName, int rowIndex, int colIndex, string cellFormulaString,bool isNeedCaculateNow=false)
        {
            var sheet = Workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                return;
            }
            var cell = sheet.Cells[rowIndex, colIndex];
            if (cell == null)
            {
                return;
            }
            cell.Formula = cellFormulaString;
            if (isNeedCaculateNow)
            {
                Workbook.CalculateFormula();
            }
        }

        #endregion

    }
}
