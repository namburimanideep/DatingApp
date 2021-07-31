using System;
using System.Collections.Generic;
using System.Web;
using Excel;
using System.Data;
using System.IO;
using ClosedXML.Excel;
using System.ComponentModel;
namespace GlobalNetApps.Support.Common
{
    public class ExcelMethods
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(typeof(ExcelMethods));
        TypeConversion tc = new TypeConversion();
        #region Excel      
        /// <summary>
        /// Generates excel for a list of model 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="fileName"></param>
        public void downloadexcelforModel<T>(List<T> data, string fileName)
        {
            try
            {
                downloadexcel(tc.getDataTableforModel(data), fileName);
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }
        }
        public void downloadexcelforModel<T>(List<T> data, string fileName, List<string> ColumsToBeRemoved)
        {
            try
            {
                downloadexcel(tc.getDataTableforModel(data), fileName, ColumsToBeRemoved);
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }
        }
        /// <summary>
        /// Generates excel from a data set
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="fileName"></param>
        public void downloadexcelFromDataSet(DataSet ds, string fileName)
        {
            try
            {
                using (XLWorkbook wb = new XLWorkbook())
                {
                    for (int i = 0; i < ds.Tables.Count; i++)
                    {
                        DataTable dt = ds.Tables[i];
                        if (dt.TableName == null || dt.TableName == "")
                        {
                            dt.TableName = "Sheet" + i.ToString();
                        }
                        wb.Worksheets.Add(dt);
                    }
                    downloadexcelFromWorkBook(wb, fileName);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }
        }
        /// <summary>
        /// Downloads excel from the work book
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="fileName"></param>
        public void downloadexcelFromWorkBook(XLWorkbook wb, string fileName)
        {
            try
            {
                string fileFullName = System.Web.HttpContext.Current.Server.UrlEncode(fileName + "_" + DateTime.Now.ToString("dd/MM/yyyy") + "_" + DateTime.Now.ToString("HH:mm:ss") + ".xlsx");
                MemoryStream stream = GetStream(wb);
                HttpContext.Current.Response.Clear();
                HttpContext.Current.Response.Buffer = false;
                HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=" + fileFullName);
                HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
                HttpContext.Current.Response.BinaryWrite(stream.ToArray());
                HttpContext.Current.Response.End();
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }
        }
        /// <summary>
        /// Creates workbook from data table
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="fileName"></param>
        public void downloadexcel(DataTable dt, string fileName)
        {
            try
            {
                dt.TableName = fileName;
                using (XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(dt);
                    downloadexcelFromWorkBook(wb, fileName);
                }
            }
            catch (System.Exception excpt)
            {
                throw (excpt);
            }
        }
        public void downloadexcel(DataTable dt, string fileName, List<string> ColumsToBeRemoved)
        {
            try
            {
                dt.TableName = fileName;
                if (ColumsToBeRemoved.Count > 0)
                {
                    for (int i = 0; i < ColumsToBeRemoved.Count; i++)
                    {
                        dt.Columns.Remove(ColumsToBeRemoved[i]);
                    }
                }
                using (XLWorkbook wb = new XLWorkbook())
                {
                    if (dt.Rows.Count > 0)
                    {
                        wb.Worksheets.Add(dt);
                        downloadexcelFromWorkBook(wb, fileName);
                    }
                }
            }
            catch (System.Exception excpt)
            {
                throw (excpt);
            }
        }

        /// <summary>
        /// Renders excel to browser
        /// </summary>
        /// <param name="excelWorkbook"></param>
        /// <returns></returns>
        public MemoryStream GetStream(XLWorkbook excelWorkbook)
        {
            MemoryStream fs = new MemoryStream();
            try
            {
                excelWorkbook.SaveAs(fs);
                fs.Position = 0;
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
                throw (ex);
            }
            return fs;
        }
        /// <summary>
        /// Reads data from excel file in a specified path and returnsas a data table
        /// </summary>
        /// <param name="rootFolderName"></param>
        /// <param name="FileNameWithExtension"></param>
        /// <returns></returns>
        public DataTable readDataFromExcel(string rootFolderName, string FileNameWithExtension, int sheetIndex)
        {
            DataSet dsData = new DataSet();
            try
            {
                FileStream dataStream = System.IO.File.Open(System.Web.HttpContext.Current.Server.MapPath("~/" + rootFolderName) + "\\" + FileNameWithExtension, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(dataStream);
                dsData = excelReader.AsDataSet();
                foreach (DataColumn column in dsData.Tables[sheetIndex].Columns)
                {
                    string cName = dsData.Tables[sheetIndex].Rows[sheetIndex][column.ColumnName].ToString();
                    if (!dsData.Tables[sheetIndex].Columns.Contains(cName) && cName != "")
                    {
                        column.ColumnName = cName.Trim();
                    }
                }
                dsData.Tables[sheetIndex].AcceptChanges();
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
                throw (ex);
            }
            return dsData.Tables[sheetIndex];
        }
        #endregion
    }
}