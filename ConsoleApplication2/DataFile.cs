using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2
{
    public class UnamePass
    {
        public string Source { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
    }

    public class CardInfo
    {
        public string Source { get; set; }
        public string CardNo { get; set; }
        public string IFSCCODE { get; set; }
        public string Validthrough { get; set; }
        public string ValidFrom { get; set; }
        public string NameOnCard { get; set; }
        public string ThreeDSecureCode { get; set; }
        public string CVV { get; set; }
        public string Notes { get; set; }
    }

    public class Mydata
    {
        public List<UnamePass> Unamepass { get; set; }
        public List<CardInfo> Cardinfo { get; set; }
    }

    public class FileData
    {
        public Mydata Mydata { get; set; }
    }

}
