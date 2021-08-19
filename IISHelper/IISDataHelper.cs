using System;
using System.Linq;
using System.Xml.Linq;
using System.DirectoryServices;
using System.Collections.Generic;
using DTO;

namespace IISHelper {
    public class IISDataHelper {

        #region Generat XML Sehema        

        public static XDocument ReadSettings(string ip, string uid, string pwd, Action<string> progressCallback) {
            string path = String.Format("IIS://{0}/W3SVC", ip);
            DirectoryEntry w3svc =
                string.IsNullOrEmpty(uid) ?
                new DirectoryEntry(path) :
                new DirectoryEntry(path, uid, pwd);
            XDocument xd = XDocument.Parse("<root />");
            pools.Clear();
            exploreTree(w3svc, xd.Root, progressCallback);
            return xd;
        }

        static Dictionary<string, string> pools =
            new Dictionary<string, string>();

        private static void exploreTree(DirectoryEntry de, XElement xe, Action<string> cb) {
            foreach (DirectoryEntry childEntry in de.Children) {
                XElement childElement = new XElement(
                    childEntry.SchemaClassName,
                    new XAttribute("Name", childEntry.Name),
                    new XAttribute("Path", childEntry.Path)
                    );

                //Get properties
                XElement propNode = new XElement("Properties");
                foreach (PropertyValueCollection pv in childEntry.Properties) {
                    //Array
                    if (pv.Value != null && pv.Value.GetType().IsArray) {
                        XElement propCol = new XElement(pv.PropertyName);
                        foreach (object obj in pv.Value as object[]) {
                            string v = Convert.ToString(obj);
                            //Set ASP.NET version
                            if (pools.Count == 0 && pv.PropertyName == "ScriptMaps" && v.StartsWith(".aspx,")) {
                                string aspNetVer = v.Split(',')[1].Split('\\') .Single(o => o.StartsWith("v"));
                                childElement.Add(new XAttribute("AspNetVer", aspNetVer));
                            }
                            propCol.Add(new XElement("Entry", v));
                        }
                        propNode.Add(propCol);
                    } else {
                        string v = Convert.ToString(pv.Value);
                        propNode.Add(new XElement(pv.PropertyName, v));
                        //Set home directory
                        if (pv.PropertyName == "Path") {
                            childElement.Add(new XAttribute("HomeDir", v));
                        } else if (pools.Count > 0 &&
                            pv.PropertyName == "AppPoolId" && pools.ContainsKey(v)) {
                            //Try to find the runtime version
                            childElement.Add(
                                new XAttribute("AspNetVer", pools[v]));
                        } else if (pools.Count > 0 &&
                            pv.PropertyName == "Enable32BitAppOnWin64" && pools.ContainsKey(v)) {
                            //Try to find the runtime version
                            childElement.Add(
                                new XAttribute("Enable32bit", pools[v]));
                        }

                    }
                }
                //For IIS 7, use AppPool to decide ASP.NET runtime version, 新增兼容IIS6.0
                if (childEntry.SchemaClassName == "IIsApplicationPool" || childEntry.SchemaClassName == "IIsApplicationPools") {
                    XElement runtimeVer = propNode.Element("ManagedRuntimeVersion");

                    
                    string ver = runtimeVer != null ? runtimeVer.Value : "v2.0";
                    pools.Add(childEntry.Name, ver);

                   
                } 
                childElement.Add(propNode);

                xe.Add(childElement);
                cb?.Invoke(childEntry.Name);
                exploreTree(childEntry, childElement, cb);
            }
        }

        #endregion


        /// <summary>
        /// 取得集區資訊
        /// </summary>
        /// <param name="xd"></param>
        /// <returns></returns>
        public static IEnumerable<APPoolModel> GetAppPoolData(XDocument xd) {
            return from E in xd.Descendants("IIsApplicationPool")
            select new APPoolModel {
                PoolName = E.Attribute("Name").Value,
                NetVersion = E.Element("Properties").Element("ManagedRuntimeVersion") != null ? E.Element("Properties").Element("ManagedRuntimeVersion").Value : "",
                Enable32Bit = E.Element("Properties").Element("Enable32BitAppOnWin64") != null ? (bool.Parse(E.Element("Properties").Element("Enable32BitAppOnWin64").Value) == true ? "32" : "64") : ""
            };
        }

        /// <summary>
        /// 取得IIS設定站台資訊
        /// </summary>
        /// <param name="xd"></param>
        /// <returns></returns>
        public static IEnumerable<IISWebSiteModel> GetIISMap(XDocument xd) {
            return from Q in xd.Root.Descendants()
                   where Q.Attribute("HomeDir") != null &&
                         Q.Attribute("AspNetVer") != null
                   select new IISWebSiteModel {
                       AspNetVer = Q.Attribute("AspNetVer").Value,
                       SiteName = Q.Attribute("Name").Value,
                       HomeDir = Q.Attribute("HomeDir").Value,
                       URL = Q.Attribute("Path").Value,
                       Enable32Bit = "",
                       AppPoolName = Q.Element("Properties").Element("AppPoolId").Value
                   };
        }

        public  static IEnumerable<FilterModel> GetFilter(XDocument xd) {
            var _Filter = from F in xd.Root.Descendants("IIsWebServer").Descendants("IIsFilter")
                          where F.Attribute("Name").Value != null && F.Attribute("Path").Value != null
                          select new FilterModel{
                              FilterName = F.Attribute("Name").Value,
                              URLPath = F.Attribute("Path").Value.Substring(0, F.Attribute("Path").Value.IndexOf("filters") - 1)
                          };

            return _Filter;
        }

        /// <summary>
        /// 取得網站支援服務之程式
        /// </summary>
        /// <param name="xd"></param>
        /// <param name="SitePath"></param>
        /// <returns></returns>
        public static List<string> GetExecutableProgram(XDocument xd, string SitePath) {
            List<string> resultVal = new List<string>();
            var _Entry = from E in xd.Root.Descendants()
                      where E.Attribute("HomeDir") != null &&
                            E.Attribute("AspNetVer") != null
                            && E.Attribute("Path").Value.IndexOf(SitePath) >= 0
                      select E.Element("Properties").Element("ScriptMaps").Descendants("Entry").ToList();

            foreach (var item in _Entry) {
                foreach (var xm in item.Select(e => e.Element("Entry"))) {
                    if (xm.Value.IndexOf("aspnet_isapi.dll") >= 0 && !resultVal.Contains("ASP.NET")) {
                        resultVal.Add("ASP.NET");
                        continue;
                    } else if (xm.Value.IndexOf("asp.dll") >= 0 && !resultVal.Contains("ASP")) {
                        resultVal.Add("ASP");
                        continue;
                    } else if (xm.Value.IndexOf("php4isapi.dll") >= 0 && !resultVal.Contains("PHP")) {
                        resultVal.Add("PHP");
                        continue;
                    }
                }
            }

            return resultVal;

        }

    }
}