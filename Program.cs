using System;
using System.Collections.Generic;
using System.Net;
using System.Data;
using System.ComponentModel;
using Newtonsoft.Json;
using System.IO;
using OfficeOpenXml;

namespace PSD_Data_Retrive
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                List<PSD> UserList = new List<PSD>();
                string USDAUrl = System.Configuration.ConfigurationManager.AppSettings["USDAUrl"].ToString();
                string API_KEY = System.Configuration.ConfigurationManager.AppSettings["API_KEY"].ToString();
                string commodityCode = "0430000";// "0574000";
                int marketYear = 2018;
                using (var client=new System.Net.WebClient())
                {
                    var myClient = new WebClient();
                    myClient.Headers.Add("API_KEY",API_KEY);
                    //myClient.TimeOut=9000000;
                    //Daft Timeout period myClient.Headers[HttpRequestHeader.Contenttype]="application/json";
                    myClient.BaseAddress = USDAUrl + "api/CommodityData/GetCommodityDataByYear?CommodityCode=" + commodityCode + "&marketYear=" + marketYear;
                    //WebResponse response=;
                    var results = myClient.DownloadString(myClient.BaseAddress);
                    //object JsonConvert = null;
                    //updateStatus=NewtonSoft.Json.JsonConvert.DeserializeObject<bool>(Encoding.UTF8.GetString(result));
                    UserList = JsonConvert.DeserializeObject<List<PSD>>(results);
                
                    DataTable dt = ExtensionClass.ToDataTable<PSD>(UserList);
                    Program obj = new Program();
                    obj.GenerateExcel(dt,"PSD_DATA",@"D:\PSD_DATA","PSD_DATA_File");
                  
                }
            }
            catch(Exception ex)
            {
                throw;
            }
        }
        public bool GenerateExcel(DataTable Data, string WorkSheetName, string FilePath, string WorkFileName)
        {
            string fname = "";
            fname = WorkFileName + "_" + Convert.ToString(DateTime.Now.ToString("ddMMyyHHmmss")) + ".csv";
            string path1 = FilePath;
            if (!Directory.Exists(path1))
            {
                Directory.CreateDirectory(path1);
            }
            string filepath = System.IO.Path.Combine(path1, fname);
            using (FileStream fs = new FileStream(filepath, FileMode.CreateNew))
            {
                using (ExcelPackage objExcelPackage = new ExcelPackage(fs))
                {
                    objExcelPackage.Workbook.Worksheets.Add(WorkSheetName).Cells["A1"].LoadFromDataTable(Data, true);
                    ExcelWorkbook workBook = objExcelPackage.Workbook;
                    ExcelWorksheet worksheet1 = workBook.Worksheets[WorkSheetName];
                    worksheet1.Cells["A2:" + Char.ConvertFromUtf32(Data.Columns.Count + 64) + "2"].AutoFitColumns();
                    //worksheet1.Cells["A1:" + Char.ConvertFromUtf32(Data.Columns.Count + 64) + "1"].AutoFitColumns();
                    //worksheet1.Cells["A2:" + Char.ConvertFromUtf32(Data.Columns.Count + 64) + "1"].Style.Font.Bold = true;
                    objExcelPackage.Save();
                    objExcelPackage.Dispose();
                }
            }

            return true;

        }
    }
    
    public static class ExtensionClass
    {
        public static DataTable ToDataTable<T>(this IList<T> data)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor prop = props[i];
                table.Columns.Add(prop.Name, prop.PropertyType);
            }
                object[] values = new object[props.Count];
                foreach(T Item in data)
                {
                    for (int j=0;j<values.Length;j++)
                    {
                        values[j] = props[j].GetValue(Item);
                    }
                    table.Rows.Add(values);
                }
                return table;
            }
        }

    }
    public class PSD
    { 
        public string CommodityCode { get; set; }
        public string CommodityDescription { get; set; }
        public string CountryCode { get; set; }
        public string CountryName { get; set; }
        public string  MarketYear { get; set; }
        public string CalendarYear { get; set; }
        public string Month { get; set; }
        public Int32 AttributeId { get; set; }
        public string AttributeDescription { get; set; }
        public Int32 UnitId { get; set; }
        public string UnitDescription { get; set; }
        public decimal Value { get; set; }
    }

