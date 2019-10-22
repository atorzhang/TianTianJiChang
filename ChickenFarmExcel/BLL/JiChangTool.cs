using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ChickenFarmExcel.BLL
{
    public class JiChangTool
    {

        public static string CheckServer()
        {
            var checkResult = "远程校验失败，请检查网络或者联系所有人询问授权情况";
            var webClient = new WebClient { Encoding = Encoding.UTF8 };
            try
            {
                var checkResultStr = webClient.DownloadString("http://sn.ae100.top/check2019.txt");
                if (checkResultStr == "{check:1}")
                {
                    checkResult = "";
                }
            }
            catch (Exception ex)
            {
            }
            return checkResult;
        }

        /// <summary>
        /// 获取项目时间差
        /// </summary>
        /// <param name="finalData"></param>
        /// <param name="lstFinalData"></param>
        /// <returns></returns>
        public static TimeSpan GetReturnTimeSubLastOut2(int index, List<FinalData> lstFinalData, ref DateTime lastOut2)
        {
            var item = lstFinalData[index];
            for (int i = index; i < lstFinalData.Count; i++)
            {
                if (i < lstFinalData.Count - 1)
                {
                    if (lstFinalData[i + 1].CarNo != item.CarNo)
                    {
                        //说明为该车牌号的最后一条数据
                        lastOut2 = (DateTime)lstFinalData[i].Out2;
                        return (DateTime)lstFinalData[i].ReturnTime - (DateTime)lstFinalData[i].Out2;

                    }
                    else
                    {
                        continue;
                    }
                }
                else
                {
                    //最后一个了
                    lastOut2 = (DateTime)lstFinalData[i].Out2;
                    return (DateTime)lstFinalData[i].ReturnTime - (DateTime)lstFinalData[i].Out2;
                }
            }
            lastOut2 = (DateTime)item.Out2;
            return (DateTime)item.ReturnTime - (DateTime)item.Out2;
        }
    }
}
