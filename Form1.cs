using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Threading;
using System.Runtime.InteropServices;


namespace checkdataCollect
{
   
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        private bool threadStart = false;
        Microsoft.Office.Interop.Excel.Application app;
        Workbooks wbks;
        Workbook r_wbk;
        Workbook w_wbk;

        peopleSet ps = new peopleSet(2);
        Dictionary<KeyValuePair<String, String>, int> itemDic = new Dictionary<KeyValuePair<String, String>, int>();     //检查项目对应的列的字典
        string fileName="";
        string fileWay_Source;
        string fileWay_Destination;

        #region Excel相关变量定义及操作


        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        private void openExcel(String filename)
        {
            app.Visible = false;
            this.wbks = app.Workbooks;
            r_wbk = wbks.Add(filename);
        }

        private void createExcel()
        {
            app.AlertBeforeOverwriting = false;         // 屏蔽系统的alert
            w_wbk = wbks.Add(true);
            w_wbk.SaveAs(@System.Environment.CurrentDirectory + "\\" + fileName + ".xls", 56, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        }

        private void CloseExcel(Workbook wb)
        {
            wb.Save();
            wb.Close();
        }

        // 结束 Excel 进程
        public static void KillExcel(Microsoft.Office.Interop.Excel.Application excel)
        {
            IntPtr t = new IntPtr(excel.Hwnd);
            int k = 0;
            GetWindowThreadProcessId(t, out k);
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
            p.Kill();
        }
        #endregion

        //进度条更新函数
        private delegate void SetPos(int ipos);
        private void SetTextMessage(int ipos)
        {
            if (this.InvokeRequired)
            {
                SetPos setpos = new SetPos(SetTextMessage);
                this.Invoke(setpos, new object[] { ipos });
            }
            else
            {
                this.label2.Text = ipos.ToString() + "/100";
                this.progressBar1.Value = Convert.ToInt32(ipos);
            }
        }

        //选择源文件
        private void selectfolderbtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            openFileDialog.Filter = "Excel文件|*.xls*";
            openFileDialog.RestoreDirectory = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.txtboxPath.Text = openFileDialog.FileName;
                fileName =System.IO.Path.GetFileName(openFileDialog.FileName)+"(汇总结果)";
            }
        }

        //读取一个工作簿的检查数据，并存进peoples的集合中
        private peopleSet collectdata(string fileWay)
        {
            peopleSet ps = new peopleSet(2);                            //people的数据集合

            openExcel(fileWay);
     
            String curPeopleID;                                         //存储当前的检测用户

            Sheets r_shs = r_wbk.Sheets;                                //获取excel工作簿的所有工作表
            _Worksheet r_wsh = (_Worksheet)r_shs.get_Item(1);           //获取第一个工作表
            long i=1;                                                   //记录行的位置
            if (Convert.ToString(((Range)r_wsh.Cells[i, ps.pItemData_column]).Text) == "")               //第一行为空
                i = 2;
            int j=ps.itemBegin_column;                                     //数据项的起始位置

            peopledata curpeople ;                                      // 存取当前的体检人员的数据
            KeyValuePair<String,String> itemset_item;

            long allrow = r_wsh.UsedRange.CurrentRegion.Rows.Count;

            this.label3.Text = "正在读取数据,请稍后……";
            while(i <= allrow)
            {
                curPeopleID = Convert.ToString(((Range)r_wsh.Cells[i, ps.pID_column]).Text);
                if (ps.peopledic.ContainsKey(curPeopleID))//该用户ID已经记录
                {
                     itemset_item  = new KeyValuePair<String,String>(Convert.ToString(((Range)r_wsh.Cells[i,ps.pItemSet_column]).Text),Convert.ToString(((Range)r_wsh.Cells[i,ps.pItem_column]).Text));
                     if (Convert.ToString(((Range)r_wsh.Cells[i, ps.pItemData_column]).Text) == "NULL")
                     {
                         ps.peopledic[curPeopleID].element.Add(itemset_item, Convert.ToString(((Range)r_wsh.Cells[i, ps.pItemData_column + 1]).Text));
                     }
                     else
                     {
                         ps.peopledic[curPeopleID].element.Add(itemset_item, Convert.ToString(((Range)r_wsh.Cells[i, ps.pItemData_column]).Text));
                     }

                }
                else //该用户没记录
                {                        
                    curpeople = new peopledata();

                    //添加检查项目信息
                    itemset_item = new KeyValuePair<String, String>(Convert.ToString(((Range)r_wsh.Cells[i, ps.pItemSet_column]).Text), Convert.ToString(((Range)r_wsh.Cells[i, ps.pItem_column]).Text));
                    curpeople.name = Convert.ToString(((Range)r_wsh.Cells[i, ps.PNname_column]).Text);
                    curpeople.sex = Convert.ToString(((Range)r_wsh.Cells[i, ps.pSex_column]).Text);
                    curpeople.age = Convert.ToString(((Range)r_wsh.Cells[i, ps.pAge_column]).Text);
                    if (Convert.ToString(((Range)r_wsh.Cells[i, ps.pItemData_column]).Text) == "NULL")
                    {
                        curpeople.element.Add(itemset_item, Convert.ToString(((Range)r_wsh.Cells[i, ps.pItemData_column+1]).Text));     //增加检查项目  
                    }
                    else
                    {
                        curpeople.element.Add(itemset_item, Convert.ToString(((Range)r_wsh.Cells[i,ps.pItemData_column]).Text));     //增加检查项目  
                    }
                    ps.peopledic.Add(curPeopleID, curpeople);

                }
                if(!itemDic.ContainsKey(itemset_item))//不包含这个项目
                {
                    itemDic.Add(itemset_item, j);
                    ++j;
                }
                if (i < allrow)
                {
                    SetTextMessage((int)(i * 100 / allrow));
                }
                else 
                {
                    Thread.Sleep(100);
                    SetTextMessage(0);
                }
                ++i;
            }
            app.AlertBeforeOverwriting = false;         // 屏蔽系统的alert
            r_wbk.Close();
            return ps;
        }

        //把数据写入一个工作簿
        private void writedata(peopleSet ps)                                           //people的数据集合)
        {
            createExcel();
            Sheets w_shs = w_wbk.Sheets;                                               //获取excel工作簿的所有工作表
            _Worksheet w_wsh = (_Worksheet)w_shs.get_Item(1);                          //获取第一个工作表
            long i = 2;
            int lineheader = 1;//定义列头开始行
            int n = ps.peopledic.Count;
            this.label3.Text = "正在写入汇总结果，请稍后……";

            w_wsh.Cells[lineheader, ps.pID_column] = "体检号";
            w_wsh.Cells[lineheader, ps.PNname_column] = "姓名";
            w_wsh.Cells[lineheader, ps.pSex_column] = "性别";
            w_wsh.Cells[lineheader, ps.pAge_column] = "年龄";
            foreach (KeyValuePair<KeyValuePair<String, String>, int> kp in itemDic)
            {
                w_wsh.Cells[lineheader, kp.Value] = kp.Key.Value;       
            }

            foreach (KeyValuePair<String,peopledata> p in ps.peopledic)
            { 
                w_wsh.Cells[i,ps.pID_column]=p.Key;
                w_wsh.Cells[i,ps.PNname_column] = p.Value.name;
                w_wsh.Cells[i, ps.pSex_column] = p.Value.sex;
                w_wsh.Cells[i, ps.pAge_column] = p.Value.age;
                int itcolumn;
                foreach (KeyValuePair<KeyValuePair<String, String>, String> it in p.Value.element)
                {
                    itcolumn = itemDic[it.Key];             //获取行号
                    w_wsh.Cells[i, itcolumn] = it.Value;    //写入相应的值
                }
                SetTextMessage((int)(100 * (i -1)/ n));
                ++i;
            }
            CloseExcel(w_wbk);
            
        }

        private void collectbtn_Click(object sender, EventArgs e)
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            this.collectbtn.Enabled = false;
            if (txtboxPath.Text != "")
            {
                Thread fThread = new Thread(new ThreadStart(solve));        //开辟一个新的线
                fThread.IsBackground = true;
                fThread.Start();

                while (true)
                {
                    if (threadStart == true)
                    {
                        fThread.Abort();
                        threadStart = false;
                        itemDic.Clear();
                        this.collectbtn.Enabled = true;
                        break;

                    }
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            else
            {
                MessageBox.Show("请点击浏览按钮选择汇总文件所在的路径！");
                this.collectbtn.Enabled = true;
            }
        }

        private void solve()
        {
            peopleSet ps;                                     //people的数据集合
            DateTime startTime = DateTime.Now;
            DateTime endTime;
            ps = collectdata(txtboxPath.Text);      //读取数据
            writedata(ps);        //写数据
            label3.Text = "完成";
            endTime = DateTime.Now;
            TimeSpan costTime = endTime.Subtract(startTime);
            MessageBox.Show("处理完成"+"\n"+"本次处理花费"+costTime.Minutes.ToString()+"分钟");
            app.Quit();         //退出Excel
            KillExcel(app);
            app = null;
            threadStart = true;

        }

        private void MergeData(peopleSet ps_D, peopleSet ps_S)
        {
            foreach (KeyValuePair<string,peopledata> kvp in ps_S.peopledic)
            {
                if (ps_D.peopledic.ContainsKey(kvp.Key))
                {
                    foreach (KeyValuePair<KeyValuePair<string, string>, string> k in kvp.Value.element)
                    {
                        if (ps_D.peopledic[kvp.Key].element.ContainsKey(k.Key))
                        {
                            ps_D.peopledic[kvp.Key].element[k.Key] = k.Value;
                        }
                    }
                }
            }
        }

        private void AddData(string fileWayDestination, string fileWaySource)
        {
            peopleSet ps_Destination;                                   //people_Destination的数据集合
            peopleSet ps_Source;                                        //people_Source的数据集合

            DateTime startTime = DateTime.Now;
            DateTime endTime;

            ps_Destination = collectdata(fileWay_Destination);      //读取数据
            ps_Source = collectdata(fileWay_Source);

            MergeData(ps_Destination, ps_Source);

            writedata(ps_Destination);        //写数据
            label3.Text = "完成";
            endTime = DateTime.Now;
            TimeSpan costTime = endTime.Subtract(startTime);
            MessageBox.Show("处理完成" + "\n" + "本次处理花费" + costTime.Minutes.ToString() + "分钟");

            app.Quit();         //退出Excel
            KillExcel(app);
            app = null;
            threadStart = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

            //选中实验室
            if (radioButton1.Checked == true) 
            {
                    //ps.pID_column =1;            //体检人员的ID所在列
                    //ps.PNname_column=2;          //体检人员的名字所在列
                    //ps.pSex_column=3;            //性别所在列
                    //ps.pAge_column=4;            //年龄所在列
                    //ps.pItemSet_column=5;        //相关检查项的集合所在列
                    //ps.pItem_column=6;           //检查项所在列
                    //ps.pItemData_column=7;       //检查项的数据所在列
                    //ps.itemBegin_column = 5;     //检查项起始列
                ps.ExcelStyleChange(2);
            }
            else if (radioButton2.Checked == true)
            {
                ps.pID_column = 1;            //体检人员的ID所在列
                ps.PNname_column = 2;          //体检人员的名字所在列
                ps.pSex_column = 3;            //性别所在列
                ps.pAge_column = 4;            //年龄所在列
                ps.pItemSet_column = 7;        //相关检查项的集合所在列
                ps.pItem_column = 8;           //检查项所在列
                ps.pItemData_column = 9;       //检查项的数据所在列
                ps.itemBegin_column = 5;     //检查项起始列               
            }
            else
            {
                MessageBox.Show("请选择一种数据格式！");          
            }   


        }

        private void button_AddData_Click(object sender, EventArgs e)
        {
            fileWay_Destination = this.txtboxPath.Text;
            if (fileWay_Destination != "")
            {
                Form2 f = new Form2(fileWay_Destination);
                f.ShowDialog();
                if (f.fileWay_Source != "")
                {
                    fileWay_Source = f.fileWay_Source;
                    app = new Microsoft.Office.Interop.Excel.Application();
                    this.collectbtn.Enabled = false;
                    this.button_AddData.Enabled = false;

                    Thread fThread = new Thread(()=> AddData(fileWay_Destination,fileWay_Source));        //开辟一个新的线
                    fThread.IsBackground = true;
                    fThread.Start();

                    while (true)
                    {
                        if (threadStart == true)
                        {
                            fThread.Abort();
                            threadStart = false;
                            itemDic.Clear();
                            this.collectbtn.Enabled = true;
                            this.button_AddData.Enabled = true;
                            break;

                        }
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
            }
            else
            {
                MessageBox.Show("请点击浏览按钮选择汇总文件所在的路径！");
            }
        }
    }
}
