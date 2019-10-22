using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChickenFarmExcel
{
    public class CSVHelper
    {
        //打开csv文件
        public static bool OpenCSVFile(ref DataTable mycsvdt, string filepath)
        {
            string strpath = filepath; //csv文件的路径
            try
            {
                int intColCount = 0;
                bool blnFlag = true;

                DataColumn mydc;
                DataRow mydr;

                string strline;
                string[] aryline;
                StreamReader mysr = new StreamReader(strpath, System.Text.Encoding.Default);
                while ((strline = mysr.ReadLine()) != null)
                {
                    aryline = strline.Split(new char[] { ',' });
                    //给datatable加上列名
                    if (blnFlag)
                    {
                        blnFlag = false;
                        intColCount = aryline.Length;
                        int col = 0;
                        for (int i = 0; i < aryline.Length; i++)
                        {
                            col = i + 1;
                            if (i > 1)
                            {
                                mydc = new DataColumn("col" + col.ToString(), Type.GetType("System.DateTime"));
                            }
                            else
                            {
                                mydc = new DataColumn("col" + col.ToString(), Type.GetType("System.String"));
                            }
                            mycsvdt.Columns.Add(mydc);
                        }
                        continue;
                    }
                    //填充数据并加入到datatable中
                    mydr = mycsvdt.NewRow();
                    for (int i = 0; i < intColCount; i++)
                    {
                        if (i < 2)
                        {
                            mydr[i] = aryline[i];
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(aryline[i]))
                            {
                                mydr[i] = DateTime.Today;
                            }
                            else
                            {
                                mydr[i] = Convert.ToDateTime(aryline[i]);
                            }
                            
                        }
                    }
                    mycsvdt.Rows.Add(mydr);
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
                throw ( new Exception($"读取CSV文件中的数据出错{ex}"));
            }
        }
    }
}
