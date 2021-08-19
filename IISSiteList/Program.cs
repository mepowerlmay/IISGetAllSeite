using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using ClosedXML.Excel;
using IISHelper.Extensions;
using DTO;

namespace IISHelper {
    class Program {
        static void Main(string[] args) {

            XDocument xd =
                IISDataHelper.ReadSettings(
                    "【AD網域】", @"AD帳號【完整】", "【密碼】",  /* 記得自己改一下 */
                    (s) => {
                        //利用\r在同一列顯示目前處理進度
                        Console.Write("\rProcessing " + s.PadRight(50));
                    }
                );
            Console.WriteLine("\rDone!{0}", new String(' ', 30));

            // 產出XML檔案比對檢視XML TAG
            //xd.Save(AppDomain.CurrentDomain.BaseDirectory + "\\Site.xml");

           var q = from o in xd.Root.Descendants()
                    where o.Attribute("HomeDir") != null &&
                          o.Attribute("AspNetVer") != null
                    select o;


            var exp = IISDataHelper.GetAppPoolData(xd);

            var SitData = IISDataHelper.GetIISMap(xd);

            
            ExportExcelFile(SitData, exp);

            foreach (var o in q) {

                Console.WriteLine(
                    "ASP.NET {0} Web [{1}] at {2} Path：{3}",
                    o.Attribute("AspNetVer").Value,
                    o.Attribute("Name").Value,
                    o.Attribute("HomeDir").Value,
                    o.Attribute("Path").Value);
            }

            Console.Read();



        }


        

        private static void ExportExcelFile(IEnumerable<IISWebSiteModel> data, IEnumerable<APPoolModel> app) {

            using (XLWorkbook wb = new XLWorkbook()) {



                var _expdata = from Q in data
                               join E in app
                                 on Q.AppPoolName equals E.PoolName into Sub
                               from E in Sub.DefaultIfEmpty()
                               select new ExportSiteListModel {
                                   Name = Q.SiteName,
                                   SitePath = string.Format("{0}", Q.URL.Substring(Q.URL.ToUpper().IndexOf("ROOT") + 4)),
                                   AspNetVer = Q.AspNetVer,
                                   HomeDir = Q.HomeDir,
                                   URL = Q.URL,
                                   AppPoolName = Q.AppPoolName,
                                   Enable32Bit = string.IsNullOrEmpty(E.Enable32Bit) ? Q.Enable32Bit : E.Enable32Bit
                               };


                //一個wrokbook內至少會有一個worksheet,並將資料Insert至這個位於A1這個位置上
                var ws = wb.Worksheets.Add("SiteList", 1);


                var HeahCol = GetPropertyNames(new ExportSiteListModel());
                ws.Cell(1, 1).InsertData(HeahCol, true);

                ws.Row(1).Cells().Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF8DC");
                ws.Row(1).Cells().Style.Font.Bold = true;
                ws.Row(1).Cells().Style.Font.FontSize = 14;
                ws.Row(1).Cells().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Row(1).Cells().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    
                //注意官方文件上說明,如果是要塞入Query後的資料該資料一定要變成是data.AsEnumerable()
                //但是我查詢出來的資料剛好是IQueryable ,其中IQueryable有繼承IEnumerable 所以不需要特別寫
                
                ws.Cell(2, 1).InsertData(_expdata);


                wb.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "\\Site.xlsx");

            }
            
        }


        /// <summary>
        /// 取得顯示欄位名稱
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        /// <remarks>
        /// 以 DisplayName 為主
        /// </remarks>
        private static List<string> GetPropertyNames(ExportSiteListModel model) {

            List<string> HeaderColumns = new List<string>();

            var obj = ObjectExtensions.CloneObject(model);

           

            foreach (MemberInfo pi in obj.GetType().GetProperties()) {
                HeaderColumns.Add(pi.GetCustomAttributes(typeof(DisplayNameAttribute), true).Cast<DisplayNameAttribute>().Single().DisplayName);
            }

            return HeaderColumns;
        }

        
    }
}
