using ChickenFarmExcel.BLL;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ChickenFarmExcel
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            //MessageBox.Show("本软件版权归[赖星雨]所有,如非授权请勿擅自演示,修改和使用,否则保留追究法律责任","版权声明");
        }

        /// <summary>
        /// 浏览文件夹
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnView_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;//设置为选择文件夹
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                txtSource.Text = dialog.FileName;
                DirectoryInfo root = new DirectoryInfo(txtSource.Text);
                FileInfo[] files = root.GetFiles();

                this.txtScreen.Text += $"\r\n加载文件夹：{dialog.FileName}\r\n";

                int csvCount = 0;//个数统计
                foreach (var csvFile in files)
                {
                    if (csvFile.Name.ToLower().Contains("csv"))
                    {
                        csvCount++;
                        this.txtScreen.Text += $"加载文件【{csvFile.Name}】\r\n";
                    }
                }
                this.txtScreen.Text += $"本次加载了【{csvCount}】个csv文件...\r\n";
            }
        }

        /// <summary>
        /// 生成操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnCreate_Click(object sender, RoutedEventArgs e)
        {
            string checkResultStr = JiChangTool.CheckServer();
            if (!string.IsNullOrEmpty(checkResultStr))
            {
                MessageBox.Show(checkResultStr, "提示");
                return;
            }
            //最终整合的数据
            List<FinalData> lstFinalData = new List<FinalData>();
            var failDate = DateTime.Now.ToString("yyyyMMdd");
            //判断输入内容
            var filePath = txtSource.Text.Trim();
            if (string.IsNullOrEmpty(filePath) || filePath =="文件夹位置")
            {
                MessageBox.Show("请选择源文件路径","提示");
                this.txtScreen.Text += $"\r\n请选择源文件路径\r\n";
                return;
            }
            //生成Excel
            this.txtScreen.Text += $"工作中...\r\n";
            //this.txtScreen.Text += $"竭尽全力...>_<.....\r\n";
            //this.txtScreen.Text += $"为了静香...>_<.....\r\n";
            //this.txtScreen.Text += $"...>_<.....\r\n";
            //获取文件
            DirectoryInfo root = new DirectoryInfo(filePath);
            FileInfo[] files = root.GetFiles();
            List<DataTable> lstDataTable = new List<DataTable>();
            foreach (var csvFile in files)
            {
                DataTable dt = new DataTable();
                if (csvFile.Name.ToLower().Contains("csv"))
                {
                    var result = CSVHelper.OpenCSVFile(ref dt, csvFile.FullName);
                    if (result)
                    {
                        lstDataTable.Add(dt);
                    }
                }
            }
            //循环操作
            string TempletFileName =System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "ExcelTemp.xlsx");
            IWorkbook wk = null;
            using (FileStream fs = File.Open(TempletFileName, FileMode.Open,
            FileAccess.Read, FileShare.ReadWrite))
            {
                //把xlsx文件读入workbook变量里，之后就可以关闭了
                wk = new XSSFWorkbook(fs);
                fs.Close();
            }
            foreach (var dataTable in lstDataTable)
            {
                //获取所有车牌号集合
                List<string> lstCarNo = new List<string>();
                //获取所有地标集合（xxx鸡场，加工xxx,圣农车队）
                List<string> lstAddr = new List<string>();
                var ChickenFarm = "";//养鸡场名称
                foreach (DataRow row in dataTable.Rows)
                {
                    var carNo = row[0]+"";
                    var addr = row[1] + "";
                    if (!lstCarNo.Contains(carNo))
                    {
                        if(carNo !="车牌号码")
                            lstCarNo.Add(carNo);
                    }
                    if (!lstAddr.Contains(addr))
                    {
                        if (addr != "地标名称")
                            lstAddr.Add(addr);
                        if(string.IsNullOrEmpty(ChickenFarm) && addr.Contains("鸡场"))
                        {
                            ChickenFarm = addr;
                        }
                    }
                }

                this.txtScreen.Text += $"...…^_<.....\r\n";
                //循环找到这个车牌所有的数据
                InOutRecode inOutRecode1 = new InOutRecode() { PlantName = "加工一厂"};
                InOutRecode inOutRecode2 = new InOutRecode() { PlantName = "加工二厂" };

                foreach (var itemCarNo in lstCarNo)
                {
                    //col1=车牌号码
                    //col2=地标名称
                    //col3=进入时间DatetimeL类型
                    //col4=离开时间DatetimeL类型
                    //根据车牌号筛选，根据入场时间排序得到数据
                    var thisCarNoRows = dataTable.Select($"col1='{itemCarNo}'"," col3 asc");

                    //计算离开时间和回来时间，都为圣农车队
                    DateTime? DepartureTime = null;//离开时间
                    DateTime? ReturnTime = null;//返回时间
                    #region 计算离开时间
                    for (int i = 0; i < thisCarNoRows.Length; i++)
                    {
                        if (i + 1 < thisCarNoRows.Length)
                        {
                            if (thisCarNoRows[i]["col2"].ToString() == "圣农车队")
                            {
                                if (thisCarNoRows[i + 1]["col2"].ToString() != "圣农车队")
                                {
                                    DepartureTime = ((DateTime)thisCarNoRows[i]["col4"]);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            DepartureTime = ((DateTime)thisCarNoRows[i]["col4"]);
                        }
                    }
                    #endregion

                    #region 计算返回时间
                    for (int i = thisCarNoRows.Length - 1; i >= 0; i--)
                    {
                        if (i - 1 > 0)
                        {
                            if (thisCarNoRows[i]["col2"].ToString() == "圣农车队")
                            {
                                if (thisCarNoRows[i - 1]["col2"].ToString() != "圣农车队")
                                {
                                    ReturnTime = ((DateTime)thisCarNoRows[i]["col3"]);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            ReturnTime = ((DateTime)thisCarNoRows[i]["col3"]);
                        }
                    }
                    #endregion

                    //循环添加数据吧，路线车队-》鸡场-》加工厂-》鸡场-》加工厂。。。。。-》车队
                    var selectThisCarNoRows = thisCarNoRows.Where(o => o["col2"] + "" != "圣农车队").ToList();//排除圣农车队和加工二厂的数据
                    FinalData finalData = new FinalData() {
                        CarNo = itemCarNo,
                        DepartureTime = DepartureTime,
                        ReturnTime = ReturnTime,
                        Slaughterhouse = "加工一厂",
                        ChickenFarm = ChickenFarm,
                        Come2 = DateTime.Today,//默认值
                        Out2 = DateTime.Today//默认值
                    };
                    for (int i = 0; i < selectThisCarNoRows.Count; i++)
                    {
                        if(i+1 < selectThisCarNoRows.Count)
                        {
                            //进入鸡场时间记录
                            if (selectThisCarNoRows[i]["col2"].ToString().Contains("鸡场") &&(i == 0 || !selectThisCarNoRows[i-1]["col2"].ToString().Contains("鸡场")))
                            {
                                finalData.Come1 = (DateTime)selectThisCarNoRows[i]["col3"];
                            }
                            //出去鸡场时间记录
                            if (selectThisCarNoRows[i]["col2"].ToString().Contains("鸡场") && (i == selectThisCarNoRows.Count-1 || !selectThisCarNoRows[i + 1]["col2"].ToString().Contains("鸡场")))
                            {
                                finalData.Out1 = (DateTime)selectThisCarNoRows[i]["col4"];
                            }
                            //进入加工厂记录
                            #region 之前记录
                            //if(i > 0)
                            //{
                            //    if (selectThisCarNoRows[i]["col2"].ToString().Contains("加工") && (!selectThisCarNoRows[i - 1]["col2"].ToString().Contains("加工")))
                            //    {
                            //        finalData.Come2 = (DateTime)selectThisCarNoRows[i]["col3"];
                            //    }
                            //} 
                            #endregion

                            #region 改造后记录，把加工厂数据全部存储起来
                            if (selectThisCarNoRows[i]["col2"].ToString().Contains("加工一厂"))
                            {
                                inOutRecode1.InTime.Add((DateTime)selectThisCarNoRows[i]["col3"]);
                                inOutRecode1.OutTime.Add((DateTime)selectThisCarNoRows[i]["col4"]);
                            }
                            if (selectThisCarNoRows[i]["col2"].ToString().Contains("加工二厂"))
                            {
                                inOutRecode2.InTime.Add((DateTime)selectThisCarNoRows[i]["col3"]);
                                inOutRecode2.OutTime.Add((DateTime)selectThisCarNoRows[i]["col4"]);
                            } 
                            #endregion

                            //出去加工厂记录，不包含最后一条
                            if (selectThisCarNoRows[i]["col2"].ToString().Contains("加工") && (!selectThisCarNoRows[i + 1]["col2"].ToString().Contains("加工")))
                            {
                                #region 原来记录
                                //finalData.Out2 = (DateTime)selectThisCarNoRows[i]["col4"];
                                ////把记录添加了再重新New一个FinalData对象
                                //lstFinalData.Add(finalData);
                                //finalData = new FinalData
                                //{
                                //    CarNo = itemCarNo,
                                //    DepartureTime = DepartureTime,
                                //    ReturnTime = ReturnTime,
                                //    Slaughterhouse = "加工一厂",
                                //    ChickenFarm = ChickenFarm
                                //};
                                #endregion

                                //改造记录
                                #region 进出时间计算最大值
                                //1厂进出时间
                                var inOutRecode1Come = DateTime.Today;
                                var inOutRecode1Out = DateTime.Today;
                                if (inOutRecode1.InTime.Count > 0)
                                {
                                    inOutRecode1Come = inOutRecode1.InTime.OrderBy(o => o).FirstOrDefault();
                                    inOutRecode1Out = inOutRecode1.OutTime.OrderByDescending(o => o).FirstOrDefault();
                                }
                                //2场进出时间
                                var inOutRecode2Come = DateTime.Today;
                                var inOutRecode2Out = DateTime.Today;
                                if (inOutRecode2.InTime.Count > 0)
                                {
                                    inOutRecode2Come = inOutRecode2.InTime.OrderBy(o => o).FirstOrDefault();
                                    inOutRecode2Out = inOutRecode2.OutTime.OrderByDescending(o => o).FirstOrDefault();
                                } 
                                #endregion

                                //说明有2厂的数据
                                if(inOutRecode2Come != DateTime.Today && inOutRecode2Out != DateTime.Today)
                                {
                                    //有2厂数据且大于30分钟，就认为是2厂的数据
                                    if(inOutRecode2Out - inOutRecode2Come > TimeSpan.FromMinutes(20))
                                    {
                                        finalData.Out2 = inOutRecode2Out;
                                        finalData.Come2 = inOutRecode2Come;
                                        finalData.Slaughterhouse = "加工二厂";
                                        lstFinalData.Add(finalData);
                                        //归零
                                        finalData = new FinalData
                                        {
                                            CarNo = itemCarNo,
                                            DepartureTime = DepartureTime,
                                            ReturnTime = ReturnTime,
                                            Slaughterhouse = "加工一厂",
                                            ChickenFarm = ChickenFarm,
                                            Come2 = DateTime.Today,//默认值
                                            Out2 = DateTime.Today//默认值
                                        };
                                        inOutRecode1.InTime = new List<DateTime>();
                                        inOutRecode1.OutTime = new List<DateTime>();
                                        inOutRecode2.InTime = new List<DateTime>();
                                        inOutRecode2.OutTime = new List<DateTime>();
                                        continue;
                                    }
                                }
                                //2厂没数据就是1厂的东西了
                                finalData.Out2 = inOutRecode1Out;
                                finalData.Come2 = inOutRecode1Come;
                                lstFinalData.Add(finalData);
                                //归零
                                finalData = new FinalData
                                {
                                    CarNo = itemCarNo,
                                    DepartureTime = DepartureTime,
                                    ReturnTime = ReturnTime,
                                    Slaughterhouse = "加工一厂",
                                    ChickenFarm = ChickenFarm,
                                    Come2 = DateTime.Today,//默认值
                                    Out2 = DateTime.Today//默认值
                                };
                                inOutRecode1.InTime = new List<DateTime>();
                                inOutRecode1.OutTime = new List<DateTime>();
                                inOutRecode2.InTime = new List<DateTime>();
                                inOutRecode2.OutTime = new List<DateTime>();

                            }
                        }
                        else
                        {
                            //最后一条记录，如果是鸡场而不是加工厂也把记录添加了
                            //最后一条记录是鸡场的话，那就没有加工厂的进出时间
                            if (selectThisCarNoRows[i]["col2"].ToString().Contains("鸡场"))
                            {
                                if (finalData.Come1 == null)
                                {
                                    finalData.Come1 = (DateTime)selectThisCarNoRows[i]["col3"];
                                }
                                finalData.Out1 = (DateTime)selectThisCarNoRows[i]["col4"];
                            }
                            //正常的加工厂最后一条记录
                            else if (selectThisCarNoRows[i]["col2"].ToString().Contains("加工"))
                            {
                                //if (finalData.Come2 == null)
                                //{
                                //    finalData.Come2 = (DateTime)selectThisCarNoRows[i]["col3"];
                                //}
                                //finalData.Out2 = (DateTime)selectThisCarNoRows[i]["col4"];
                                #region 改造后记录，把加工厂数据全部存储起来
                                if (selectThisCarNoRows[i]["col2"].ToString().Contains("加工一厂"))
                                {
                                    inOutRecode1.InTime.Add((DateTime)selectThisCarNoRows[i]["col3"]);
                                    inOutRecode1.OutTime.Add((DateTime)selectThisCarNoRows[i]["col4"]);
                                }
                                if (selectThisCarNoRows[i]["col2"].ToString().Contains("加工二厂"))
                                {
                                    inOutRecode2.InTime.Add((DateTime)selectThisCarNoRows[i]["col3"]);
                                    inOutRecode2.OutTime.Add((DateTime)selectThisCarNoRows[i]["col4"]);
                                }
                                #endregion

                                #region 进出时间计算最大值
                                //1厂进出时间
                                var inOutRecode1Come = DateTime.Today;
                                var inOutRecode1Out = DateTime.Today;
                                if (inOutRecode1.InTime.Count > 0)
                                {
                                    inOutRecode1Come = inOutRecode1.InTime.OrderBy(o => o).FirstOrDefault();
                                    inOutRecode1Out = inOutRecode1.OutTime.OrderByDescending(o => o).FirstOrDefault();
                                }
                                //2场进出时间
                                var inOutRecode2Come = DateTime.Today;
                                var inOutRecode2Out = DateTime.Today;
                                if (inOutRecode2.InTime.Count > 0)
                                {
                                    inOutRecode2Come = inOutRecode2.InTime.OrderBy(o => o).FirstOrDefault();
                                    inOutRecode2Out = inOutRecode2.OutTime.OrderByDescending(o => o).FirstOrDefault();
                                }
                                #endregion

                                #region 计算1，2厂
                                if (inOutRecode2Come != DateTime.Today && inOutRecode2Out != DateTime.Today)
                                {
                                    //有2厂数据且大于30分钟，就认为是2厂的数据
                                    if (inOutRecode2Out - inOutRecode2Come > TimeSpan.FromMinutes(20))
                                    {
                                        finalData.Out2 = inOutRecode2Out;
                                        finalData.Come2 = inOutRecode2Come;
                                        finalData.Slaughterhouse = "加工二厂";
                                        lstFinalData.Add(finalData);
                                        //归零
                                        finalData = new FinalData
                                        {
                                            CarNo = itemCarNo,
                                            DepartureTime = DepartureTime,
                                            ReturnTime = ReturnTime,
                                            Slaughterhouse = "加工一厂",
                                            ChickenFarm = ChickenFarm,
                                            Come2 = DateTime.Today,//默认值
                                            Out2 = DateTime.Today//默认值
                                        };
                                        inOutRecode1.InTime = new List<DateTime>();
                                        inOutRecode1.OutTime = new List<DateTime>();
                                        inOutRecode2.InTime = new List<DateTime>();
                                        inOutRecode2.OutTime = new List<DateTime>();
                                        continue;
                                    }
                                }
                                //2厂没数据就是1厂的东西了
                                finalData.Out2 = inOutRecode1Out;
                                finalData.Come2 = inOutRecode1Come;
                                //归零
                                inOutRecode1.InTime = new List<DateTime>();
                                inOutRecode1.OutTime = new List<DateTime>();
                                inOutRecode2.InTime = new List<DateTime>();
                                inOutRecode2.OutTime = new List<DateTime>(); 
                                #endregion
                            }
                            lstFinalData.Add(finalData);
                        }
                    }
                }
            }
            var data = lstFinalData;
            //this.txtScreen.Text += $"...>_^....\r\n";
            this.txtScreen.Text += $"选择保存位置\r\n";
            //Thread.Sleep(3000);
            //保存对话框
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
           
            saveFileDialog1.InitialDirectory = this.txtSource.Text;
            saveFileDialog1.FileName = $"{DateTime.Now.ToString("yyyy-MM-dd")}数据.xlsx";
            saveFileDialog1.Filter = "ext files (*.xlsx)|*.xlsx|All files(*.*)|*>**";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            var dr = saveFileDialog1.ShowDialog();
            if ((bool)dr && saveFileDialog1.FileName.Length > 0)
            {
                //判断目录是否存在
                if (!Directory.Exists(System.IO.Path.GetDirectoryName(saveFileDialog1.FileName)))
                {
                    Directory.CreateDirectory(System.IO.Path.GetDirectoryName(saveFileDialog1.FileName));
                }
                var savaPath = saveFileDialog1.FileName;
                using (FileStream fs = File.Open(TempletFileName, FileMode.Open, FileAccess.Read))
                {
                    wk = new XSSFWorkbook(fs);
                    fs.Close();

                    ISheet sheet = null;
                    sheet = wk.GetSheetAt(0);

                    IRow row = null;
                    int rowCount = sheet.LastRowNum + 1;
                    int j = 1;
                    for (int i = 0; i < lstFinalData.Count; i++)
                    {
                        var item = lstFinalData[i];
                        row = sheet.CreateRow(rowCount);
                        //row.CreateCell(0).SetCellValue(j++ +"");
                        row.CreateCell(1).SetCellValue(item.CarNo);
                        if(item.DepartureTime > item.ReturnTime)
                        {
                            row.CreateCell(2).SetCellValue("无");
                        }
                        else
                        {
                            row.CreateCell(2).SetCellValue(((DateTime)item.DepartureTime).ToString("H:mm:ss"));
                        }
                        
                        //返回时间判断
                        if ((bool)this.cekJZ.IsChecked)
                        {
                            var _cell = row.CreateCell(3);
                            _cell.SetCellValue(((DateTime)item.ReturnTime).ToString("H:mm:ss"));
                            var lastOut2DateTime = (DateTime)item.Out2;
                            var timeSpanSub = BLL.JiChangTool.GetReturnTimeSubLastOut2(i, lstFinalData, ref lastOut2DateTime);
                            //在1个小时之内
                            if (timeSpanSub > TimeSpan.FromMinutes(18) && timeSpanSub <= TimeSpan.FromMinutes(60))
                            {
                                //介于18分钟到60分钟之间，变成黄色，修改时间

                                ICellStyle style = wk.CreateCellStyle();//创建单元格样式
                                style.FillForegroundColor = HSSFColor.Yellow.Index;//设置单元格样式中的样式
                                style.FillPattern = FillPattern.SolidForeground;

                                _cell.SetCellValue(lastOut2DateTime.AddMinutes(18).ToString("H:mm:ss"));
                                _cell.CellStyle = style;//为单元格设置显示样式 
                            }
                            else if (timeSpanSub > TimeSpan.FromMinutes(60))
                            {
                                //大于1个小时，变成红色
                                ICellStyle style = wk.CreateCellStyle();//创建单元格样式
                                style.FillForegroundColor = HSSFColor.Red.Index;
                                style.FillPattern = FillPattern.SolidForeground;

                                _cell.CellStyle = style;//为单元格设置显示样式 
                            }
                        }
                        else
                        {
                            row.CreateCell(3).SetCellValue(((DateTime)item.ReturnTime).ToString("H:mm:ss"));
                        }
                        
                        row.CreateCell(4).SetCellValue(item.ChickenFarm);
                        row.CreateCell(5).SetCellValue(item.Slaughterhouse);
                        if (item.Come1 != DateTime.Today && item.Come1 != null)
                        {
                            var _cell = row.CreateCell(6);
                            _cell.SetCellValue(((DateTime)item.Come1).ToString("H:mm:ss"));
                            if (i> 1)
                            {

                                if(lstFinalData[i-1].CarNo == item.CarNo)
                                {
                                    if(lstFinalData[i - 1].Out2 != null)
                                    {
                                        if((DateTime)item.Come1 -(DateTime)lstFinalData[i - 1].Out2  <= TimeSpan.FromMinutes(15))
                                        {
                                            ICellStyle style = wk.CreateCellStyle();//创建单元格样式
                                            style.FillForegroundColor = HSSFColor.Red.Index;//设置单元格样式中的样式
                                            style.FillPattern = FillPattern.SolidForeground;
                                            _cell.CellStyle = style;//为单元格设置显示样式
                                        }
                                    }
                                }
                            }
                        }
                        else
                            row.CreateCell(6).SetCellValue("无");
                        if (item.Out1 != DateTime.Today && item.Out1 != null)
                        {
                            //小于15分钟的标成黄色警告
                            var _cell = row.CreateCell(7);
                            _cell.SetCellValue(((DateTime)item.Out1).ToString("H:mm:ss"));
                            
                            if(item.Come1 != null)
                            {
                                if ((DateTime)item.Out1 - (DateTime)item.Come1 <= TimeSpan.FromMinutes(15))
                                {
                                    ICellStyle style = wk.CreateCellStyle();//创建单元格样式
                                    style.FillForegroundColor = HSSFColor.Red.Index;//设置单元格样式中的样式
                                    style.FillPattern = FillPattern.SolidForeground;
                                    _cell.CellStyle = style;//为单元格设置显示样式
                                }
                                if ((DateTime)item.Out1 - (DateTime)item.Come1 >= TimeSpan.FromHours(3))
                                {
                                    ICellStyle style = wk.CreateCellStyle();//创建单元格样式
                                    style.FillForegroundColor = HSSFColor.Red.Index;//设置单元格样式中的样式
                                    style.FillPattern = FillPattern.SolidForeground;
                                    _cell.CellStyle = style;//为单元格设置显示样式
                                }
                            }
                        }
                        else
                            row.CreateCell(7).SetCellValue("无");
                        if (item.Come2 != DateTime.Today && item.Come2 != null)
                        {
                            var _cell = row.CreateCell(8);
                            _cell.SetCellValue(((DateTime)item.Come2).ToString("H:mm:ss"));
                            if(item.Out1 != null)
                            {
                                if ((DateTime)item.Come2 - (DateTime)item.Out1 <= TimeSpan.FromMinutes(20))
                                {
                                    ICellStyle style = wk.CreateCellStyle();//创建单元格样式
                                    style.FillForegroundColor = HSSFColor.Red.Index;//设置单元格样式中的样式
                                    style.FillPattern = FillPattern.SolidForeground;
                                    _cell.CellStyle = style;//为单元格设置显示样式
                                }
                                if ((DateTime)item.Come2 - (DateTime)item.Out1 >= TimeSpan.FromMinutes(120))
                                {
                                    ICellStyle style = wk.CreateCellStyle();//创建单元格样式
                                    style.FillForegroundColor = HSSFColor.Red.Index;//设置单元格样式中的样式
                                    style.FillPattern = FillPattern.SolidForeground;
                                    _cell.CellStyle = style;//为单元格设置显示样式
                                }
                            }
                        }
                        else
                            row.CreateCell(8).SetCellValue("无");
                        if (item.Out2 != DateTime.Today && item.Out2 != null)
                        {
                            //校正按钮开启
                            if ((bool)this.cekJZ.IsChecked)
                            {
                                var _cell = row.CreateCell(9);
                                _cell.SetCellValue(((DateTime)item.Out2).ToString("H:mm:ss"));
                                if ((DateTime)item.Out2 - (DateTime)item.Come2 <= TimeSpan.FromMinutes(15))
                                {
                                    ICellStyle style = wk.CreateCellStyle();//创建单元格样式
                                    style.FillForegroundColor = HSSFColor.Red.Index;//设置单元格样式中的样式
                                    style.FillPattern = FillPattern.SolidForeground;
                                    _cell.CellStyle = style;//为单元格设置显示样式
                                }
                                if ((DateTime)item.Out2 - (DateTime)item.Come2 >= TimeSpan.FromHours(3))
                                {
                                    ICellStyle style = wk.CreateCellStyle();//创建单元格样式
                                    style.FillForegroundColor = HSSFColor.Red.Index;//设置单元格样式中的样式
                                    style.FillPattern = FillPattern.SolidForeground;
                                    _cell.CellStyle = style;//为单元格设置显示样式
                                }
                                #region 原来判断修改的是最后一次加工厂时间，注释了
                                //归来时间大于最后一次出加工厂时间18分钟，修改最后一次出加工厂时间
                                //if (i < lstFinalData.Count - 1)
                                //{

                                //    if (item.CarNo != lstFinalData[i + 1].CarNo)
                                //    {
                                //        #region 时间判断
                                //        var timeSpanSub = (DateTime)item.ReturnTime - (DateTime)item.Out2;
                                //        //在1个小时之内
                                //        if (timeSpanSub > TimeSpan.FromMinutes(18) && timeSpanSub <= TimeSpan.FromMinutes(60))
                                //        {
                                //            //介于18分钟到60分钟之间，变成黄色，修改时间

                                //            ICellStyle style = wk.CreateCellStyle();//创建单元格样式
                                //            style.FillForegroundColor = HSSFColor.Yellow.Index;//设置单元格样式中的样式
                                //            style.FillPattern = FillPattern.SolidForeground;

                                //            var _cell = row.CreateCell(9);
                                //            _cell.SetCellValue(((DateTime)item.ReturnTime).AddMinutes(-18).ToString("H:mm:ss"));
                                //            _cell.CellStyle = style;//为单元格设置显示样式 
                                //        }
                                //        else if (timeSpanSub > TimeSpan.FromMinutes(60))
                                //        {
                                //            //大于1个小时，变成红色
                                //            ICellStyle style = wk.CreateCellStyle();//创建单元格样式
                                //            style.FillForegroundColor = HSSFColor.Red.Index;
                                //            style.FillPattern = FillPattern.SolidForeground;

                                //            var _cell = row.CreateCell(9);
                                //            row.CreateCell(9).SetCellValue(((DateTime)item.Out2).ToString("H:mm:ss"));
                                //            _cell.CellStyle = style;//为单元格设置显示样式 

                                //        }
                                //        else
                                //        {
                                //            //不变
                                //            row.CreateCell(9).SetCellValue(((DateTime)item.Out2).ToString("H:mm:ss"));
                                //        }
                                //        #endregion
                                //    }
                                //    else
                                //    {
                                //        row.CreateCell(9).SetCellValue(((DateTime)item.Out2).ToString("H:mm:ss"));
                                //    } 

                                //}
                                //else if (i == lstFinalData.Count-1)
                                //{
                                //    #region 时间判断
                                //    var timeSpanSub = (DateTime)item.ReturnTime - (DateTime)item.Out2;
                                //    //在1个小时之内
                                //    if (timeSpanSub > TimeSpan.FromMinutes(18) && timeSpanSub <= TimeSpan.FromMinutes(60))
                                //    {
                                //        //介于18分钟到60分钟之间，变成黄色，修改时间

                                //        ICellStyle style = wk.CreateCellStyle();//创建单元格样式
                                //        style.FillForegroundColor = HSSFColor.Yellow.Index;//设置单元格样式中的样式
                                //        style.FillPattern = FillPattern.SolidForeground;

                                //        var _cell = row.CreateCell(9);
                                //        _cell.SetCellValue(((DateTime)item.ReturnTime).AddMinutes(-18).ToString("H:mm:ss"));
                                //        _cell.CellStyle = style;//为单元格设置显示样式 
                                //    }
                                //    else if (timeSpanSub > TimeSpan.FromMinutes(60))
                                //    {
                                //        //大于1个小时，变成红色
                                //        ICellStyle style = wk.CreateCellStyle();//创建单元格样式
                                //        style.FillForegroundColor = HSSFColor.Red.Index;
                                //        style.FillPattern = FillPattern.SolidForeground;

                                //        var _cell = row.CreateCell(9);
                                //        row.CreateCell(9).SetCellValue(((DateTime)item.Out2).ToString("H:mm:ss"));
                                //        _cell.CellStyle = style;//为单元格设置显示样式 

                                //    }
                                //    else
                                //    {
                                //        //不变
                                //        row.CreateCell(9).SetCellValue(((DateTime)item.Out2).ToString("H:mm:ss"));
                                //    }
                                //    #endregion
                                //}
                                #endregion
                            }
                            else
                            {
                                row.CreateCell(9).SetCellValue(((DateTime)item.Out2).ToString("H:mm:ss"));
                            }
                            
                        }
                        else
                            row.CreateCell(9).SetCellValue("无");
                        rowCount++;
                    }
                    sheet.ForceFormulaRecalculation = true;
                    using (FileStream filess = File.OpenWrite(savaPath))
                    {
                        wk.Write(filess);
                        this.txtScreen.Text += $"导出完成\r\n";
                        MessageBox.Show("导出完成","提示");
                        Process.Start(System.IO.Path.GetDirectoryName(savaPath));
                    }

                }
            }
        }

        /// <summary>
        /// 关于
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnAbort_Click(object sender, RoutedEventArgs e)
        {
            Abort abort = new Abort();
            abort.ShowDialog();
        }

        private void TxtScreen_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        //记录进出时间
        public class InOutRecode
        {
            public string PlantName { get; set; }
            public List<DateTime> InTime { get; set; } = new List<DateTime>();
            public List<DateTime> OutTime { get; set; } = new List<DateTime>();
        }


    }
}
