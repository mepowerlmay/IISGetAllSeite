using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace DTO {

    [Serializable]
    public class IISWebSiteModel {


        /// <summary>
        /// 網站服務名稱
        /// </summary>
        [DisplayName("網站服務名稱")]
        public string SiteName {
            get; set;
        }


        [DisplayName("URL")]
        public string URL {
            get; set;
        }


        [DisplayName(".Net版本")]
        public string AspNetVer {
            get; set;
        }


        [DisplayName("實際檔案路徑")]
        public string HomeDir {
            get; set;
        }

        [DisplayName("集區名稱")]
        public string AppPoolName {
            get; set;
        }

        string _Enable32Bit = "";

        [DisplayName("集區執行模式")]
        public string Enable32Bit {
            get {
                return _Enable32Bit;
            }
            set {
                _Enable32Bit = value;
            }
        }

        [DisplayName("延伸可執行程式")]
        public string Filters {
            get; set;
        }
    }
}
