using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CatchUp
{
    public class DataModel
    {
        public string code { get; set; }
        public List<Lucky> data { get; set; }


    }
    public class Lucky
    {
        //{"id":"56493196","sname":"ios服","rname":"c*****","info":"莫德雷德","star":"SR"}
        public int id { get; set; }
        public string sname { get; set; }
        public string rname { get; set; }
        public string info { get; set; }
        public string star { get; set; }
        public string date { get; set; }
    }
}
