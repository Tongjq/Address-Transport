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

            IRow row = null;
            List<IRow> rows = new List<IRow>();
            for (int i = 0; i < sheet.LastRowNum + 1; i++)
            {
                row = sheet.GetRow(i);
                rows.Add(row);
                if (row != null)
                {
                    if (!string.IsNullOrEmpty(row.Cells[1].ToString().Trim()))
                    { listCity.Add(row.Cells[1].ToString()); }

                    else { listCity.Add("空行"); }
                }
                else
                {
                    listCity.Add("空行");
                }
            }
            stream.Close();
            stream.Dispose();
            return GetPoint(listCity);

        }

        private bool GetPoint(List<string> listCity)
        {
            List<string> listT = new List<string>();

            HSSFWorkbook wk = new HSSFWorkbook();
            ISheet tb = wk.CreateSheet("Point");


            tb.SetColumnWidth(0, 50 * 256);
            IRow rowhead = tb.CreateRow(0);//创建首行           
            ICell cell = rowhead.CreateCell(0);//行中创建第一列
            cell.SetCellValue("Address");
            tb.SetColumnWidth(1, 40 * 256);
            ICell cell1 = rowhead.CreateCell(1);
            cell1.SetCellValue("Lng");
            tb.SetColumnWidth(2, 40 * 256);
            ICell cell2 = rowhead.CreateCell(2);
            cell2.SetCellValue("Lat");


            for (int i = 1; i < listCity.Count(); i++)
            {
                IRow row = tb.CreateRow(i);
                #region test code
                if (!string.IsNullOrEmpty(listCity[i].TrimEnd()))
                {
                    if (listCity[i] == "空行")
                    {
                        ICell cellAddress1 = row.CreateCell(0);
                        cellAddress1.SetCellValue(listCity[i]);
                        ICell lngCell1 = row.CreateCell(1);
                        lngCell1.SetCellValue("空行");
                        ICell latCell2 = row.CreateCell(2);
                        latCell2.SetCellValue("空行");
                    }
                    else
                    {

                        ICell cellAddress = row.CreateCell(0);  //创建地址单元格
                        cellAddress.SetCellValue(listCity[i]);
                        ICell lngCell = row.CreateCell(1);  //创建经度单元格
                        lngCell.SetCellValue("11");
                        ICell latCell = row.CreateCell(2);  //创建纬度单元格
                        latCell.SetCellValue("22");
                    }
                }
                #endregion


                #region request

                /* if (!string.IsNullOrEmpty(listCity[i].TrimEnd()))
                  {
                      if (listCity[i] == "空行")
                      {
                          ICell cellAddress1 = row.CreateCell(0);
                          cellAddress1.SetCellValue(listCity[i]);
                          ICell lngCell1 = row.CreateCell(1);
                          lngCell1.SetCellValue("空行");
                          ICell latCell2 = row.CreateCell(2);
                          latCell2.SetCellValue("空行");
                      }
                      else
                      {
                          string ak = "LXaG6FhzIcVFAtcoM0T4MZ0Zg78kIymV";
                          string Url = @"http://api.map.baidu.com/geocoding/v3/?address=" + listCity[i].Trim() + "&output=json&ak=" + ak;
                          HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
                          request.KeepAlive = false;
                          request.Method = "GET";
                          request.ContentType = "application/json";
                          //request.Timeout = 50000;
                          //request.ServicePoint.ConnectionLeaseTimeout = 50000;
                          //request.ServicePoint.MaxIdleTime = 50000;

                          try
                          {
                              HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                              Stream myResponseStream = response.GetResponseStream();
                              StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
                              string retString = myStreamReader.ReadToEnd();
                              myStreamReader.Close();
                              myResponseStream.Close();
                              response.Close();
                              request.Abort();
                              //var txtLocation = retString;
                              JObject obj_RawData = (JObject)(JsonConvert.DeserializeObject(retString));
                              if (obj_RawData["result"] != null)
                              {
                                  JObject obj_Result = (JObject)(JsonConvert.DeserializeObject(obj_RawData["result"].ToString()));
                                  JObject obj_Loaction = (JObject)(JsonConvert.DeserializeObject(obj_Result["location"].ToString()));
                                  string lng = obj_Loaction["lng"].ToString(); //经度值
                                  string lat = obj_Loaction["lat"].ToString(); //纬度值
                                  listT.Add(listCity[i] + "----" + "Lng:" + lng + ",Lat:" + lat);

                                  ICell cellAddress = row.CreateCell(0);  //创建地址单元格
                                  cellAddress.SetCellValue(listCity[i]);
                                  ICell lngCell = row.CreateCell(1);  //创建经度单元格
                                  lngCell.SetCellValue(lng);
                                  ICell latCell = row.CreateCell(2);  //创建纬度单元格
                                  latCell.SetCellValue(lat);
                              }
                              else
                              {
                                  listT.Add(listCity[i] + "----" + "N/A");
                                  ICell cellAddress = row.CreateCell(0);
                                  cellAddress.SetCellValue(listCity[i]);
                                  ICell lngCell = row.CreateCell(1);
                                  lngCell.SetCellValue("没有获取到");
                                  ICell latCell = row.CreateCell(2);
                                  latCell.SetCellValue("没有获取到");
                              }
                          }
                          catch
                          {
                              return false;
                          }
                      }

                  } */
                #endregion

                else
                {
                    ICell cellAddress = row.CreateCell(0);
                    cellAddress.SetCellValue(listCity[i]);
                    ICell lngCell = row.CreateCell(1);
                    lngCell.SetCellValue("空行");
                    ICell latCell = row.CreateCell(2);
                    latCell.SetCellValue("空行");
                }


            }

            var path = Server.MapPath("/UploadExcelFile") + "/AddressPoint.xls";

            using (FileStream stream = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
            {
                wk.Write(stream);
                return true;
            }

        }

    }

}