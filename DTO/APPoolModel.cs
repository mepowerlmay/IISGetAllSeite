using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DTO {

    [Serializable]
    public class APPoolModel {
        public string PoolName {
            get; set;
        }

        public string NetVersion {
            get; set;
        }


        string _Enable32Bit = "32";
        public string Enable32Bit {
            get {
                return _Enable32Bit;
            }
            set {
                _Enable32Bit = value;
            }
        }
    }
}
