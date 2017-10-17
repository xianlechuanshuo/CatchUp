using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CatchUp
{
    using System.Configuration;
    using System.Data;
    using System.IO;
    using System.Net;

    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("本程序每分钟执行一次抓包操作");
            GetLuckys();
            //每分钟更新一次抽奖名单
            System.Timers.Timer tmr = new System.Timers.Timer();
            tmr.Interval = 1000 * 60;//1min
            //tmr.Interval = 700;
            tmr.Elapsed += Tmr_Elapsed;
            tmr.AutoReset = true; //每到指定时间Elapsed事件是触发一次（false），还是一直触发（true）
            tmr.Start();
            Console.ReadLine();
        }

        static void Tmr_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            GetLuckys();
        }

        static void GetLuckys()
        {
            string url = @"http://activity.biligame.com/board/list?game_id=112&game_key=a5f36e53ab3b0c41&rows=31";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    Stream stream = response.GetResponseStream();
                    StreamReader reader = new StreamReader(stream);
                    string result = reader.ReadToEnd();
                    var list = result.Deserialize<DataModel>();

                    var data = list.data;
                    DateTime now = DateTime.Now;
                    for (int i = 0; i < data.Count; i++)
                    {
                        data[i].date = now.ToString("yyyy-MM-dd");
                    }
                    if (data != null && data.Count > 0)
                    {
                        Output(data.ToDataTable());
                    }
                }
            }
        }

        static void Output(DataTable dt)
        {
            string relativePath = ConfigurationManager.AppSettings["relativePath"];
            FileInfo fileInfo = new FileInfo(relativePath);
            string phyPath = fileInfo.FullName;

            if (!File.Exists(phyPath))
            {
                NPOIHelper.Export(dt, "", phyPath);
                Console.WriteLine(string.Format("{0} 数据量：{1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), dt.Rows.Count));

            }
            else
            {
                int totalRowCount;
                NPOIHelper.AppendForXLSX(dt, phyPath, out totalRowCount);
                Console.WriteLine(string.Format("{0} 数据量：{1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), totalRowCount));
            }
        }

    }
}
