using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Mvc;


namespace AT.WebUI.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }



        [HttpPost]
        public ActionResult UploadFile(HttpPostedFileBase file)
        {
            if (file != null)
            {
                string fileName = file.FileName;
                string filePath = Server.MapPath("~/UploadExcelFile");

                if (Directory.GetFiles(filePath).Length > 0)
                {

                    foreach (string item in Directory.GetFiles(filePath))
                    {
                        System.IO.File.Delete(item);
                    }

                }
                file.SaveAs(Path.Combine(filePath, fileName));
                //return Content("seccuess!!!!");
                return View("Translation");
            }

            return Content("No Content!");
        }



        public bool ReadExcelToTable()
        {
            var serverpath = Server.MapPath("~/UploadExcelFile");
            DirectoryInfo folder = new DirectoryInfo(serverpath);
            var path = folder.GetFiles("*.xlsx")[0].FullName;
            return ReadExcel(path);
        }


        private bool ReadExcel(string path)
        {
            List<string> listCity = new List<string>();
            Stream stream = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            try
            {
                stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                workbook = new XSSFWorkbook(stream);
                sheet = (XSSFSheet)workbook.GetSheetAt(0);
            }
            catch
            {
                stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                workbook = new HSSFWorkbook(stream);
                sheet = (HSSFSheet)workbook.GetSheetAt(0);
            }
            finally
            {
                stream.Close();
                stream.Dispose();
            }
            //ArrayList AL = new ArrayList();//创建一个集合，用来存放Excel的标题
            //int CellsCount = sheet.GetRow(0).Cells.Count;//获得这个表的第一行的列数,也就是标题行的列数
            //for (int i = 0; i < CellsCount; i++)//将标题行的每一列存储到集合中
            //{
            //    AL.Add(sheet.GetRow(0).GetCell(i).StringCellValue);
            //}

            IRow row = null;
            // ICell cell = null;
            for (int i = 0; i <=sheet.LastRowNum; i++)//从标题行一下，也就是第二行开始遍历此表
            {
                row = sheet.GetRow(i);
                if (row != null)
                {
                    if (!string.IsNullOrEmpty(row.Cells[0].ToString().Trim()))
                        listCity.Add(row.Cells[0].ToString());
                }
            }
            stream.Close();
            stream.Dispose();
            GetPoint(listCity);

            return true;
        }

        private void GetPoint(List<string> listCity)
        {
            List<string> listT = new List<string>();

            //创建工作薄
            HSSFWorkbook wk = new HSSFWorkbook();
            //创建一个名称为mySheet的表
            ISheet tb = wk.CreateSheet("Point");


            tb.SetColumnWidth(0, 50 * 256);
            IRow rowhead = tb.CreateRow(0);//创建首行           
            ICell cell = rowhead.CreateCell(0);//行中创建第一列
            cell.SetCellValue("Address");
            tb.SetColumnWidth(1, 40 * 256);
            ICell cell1 = rowhead.CreateCell(1);
            cell1.SetCellValue("lng");
            tb.SetColumnWidth(2, 40 * 256);
            ICell cell2 = rowhead.CreateCell(2);
            cell2.SetCellValue("lat");


            for (int i = 1; i < listCity.Count(); i++)
            {
                if (!string.IsNullOrEmpty(listCity[i].TrimEnd()))
                {
                    IRow row1 = tb.CreateRow(i);
                    ICell cellAddress = row1.CreateCell(0);  //创建地址单元格
                    cellAddress.SetCellValue(listCity[i]);
                    ICell lngCell = row1.CreateCell(1);  //创建经度单元格
                    lngCell.SetCellValue("11");
                    ICell latCell = row1.CreateCell(2);  //创建纬度单元格
                    latCell.SetCellValue("22");
                }


                //IRow row = tb.CreateRow(i);
                //if (!string.IsNullOrEmpty(listCity[i].TrimEnd()))
                //{
                //    string ak = "LXaG6FhzIcVFAtcoM0T4MZ0Zg78kIymV";
                //    string Url = @"http://api.map.baidu.com/geocoding/v3/?address=" + listCity[i].Trim() + "&output=json&ak=" + ak;
                //    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
                //    request.KeepAlive = false;
                //    request.Method = "GET";
                //    request.ContentType = "application/json";
                //    //request.Timeout = 50000;
                //    //request.ServicePoint.ConnectionLeaseTimeout = 50000;
                //    //request.ServicePoint.MaxIdleTime = 50000;

                //    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                //    Stream myResponseStream = response.GetResponseStream();
                //    StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
                //    string retString = myStreamReader.ReadToEnd();
                //    myStreamReader.Close();
                //    myResponseStream.Close();
                //    response.Close();
                //    request.Abort();

                //    var txtLocation = retString;


                //    JObject obj_RawData = (JObject)(JsonConvert.DeserializeObject(retString));
                //    if (obj_RawData["result"] != null)
                //    {
                //        JObject obj_Result = (JObject)(JsonConvert.DeserializeObject(obj_RawData["result"].ToString()));
                //        JObject obj_Loaction = (JObject)(JsonConvert.DeserializeObject(obj_Result["location"].ToString()));
                //        string lng = obj_Loaction["lng"].ToString(); //经度值
                //        string lat = obj_Loaction["lat"].ToString(); //纬度值
                //        listT.Add(listCity[i] + "----" + "Lng:" + lng + ",Lat:" + lat);

                //        ICell cellAddress = row.CreateCell(0);  //创建地址单元格
                //        cellAddress.SetCellValue(listCity[i]);
                //        ICell lngCell = row.CreateCell(1);  //创建经度单元格
                //        lngCell.SetCellValue(lng);
                //        ICell latCell = row.CreateCell(2);  //创建纬度单元格
                //        latCell.SetCellValue(lng);
                //    }
                //    else
                //    {
                //        listT.Add(listCity[i] + "----" + "N/A");
                //        ICell cellAddress = row.CreateCell(0);  
                //        cellAddress.SetCellValue(listCity[i]);
                //        ICell lngCell = row.CreateCell(1);  
                //        lngCell.SetCellValue("没有获取到");
                //        ICell latCell = row.CreateCell(2);  
                //        latCell.SetCellValue("没有获取到");
                //    }

                //}
            }

            var path = Server.MapPath("/UploadExcelFile")+ "/AddressPoint.xls";

            using (FileStream stream = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
            {
               
                   wk.Write(stream);
               
            }

        }

    }

}