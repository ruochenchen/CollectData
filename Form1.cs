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


        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        // 结束 Excel 进程
        public static void KillExcel(Microsoft.Office.Interop.Excel.Application excel)
        {
            IntPtr t = new IntPtr(excel.Hwnd);
            int k = 0;
            GetWindowThreadProcessId(t, out k);
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
            p.Kill();
        }









        // private object threadStart = new object();               //开启子线程的标识
        private bool threadStart = false;
        Microsoft.Office.Interop.Excel.Application app;
         Workbooks wbks;
         Workbook r_wbk;
         Workbook w_wbk;
         List<peopledata> peoples=new List<peopledata>();                   //people的数据集合
         List<String> peoplesId = new List<string>();                       //peoplesId的集合
         Dictionary<KeyValuePair<String, String>, int> itemDic = new Dictionary<KeyValuePair<String, String>, int>();     //检查项目对应的列的字典
         String fileName="";
         private delegate void SetPos(int ipos);


         //excel里面一些位置的定义
         public  int pID_column =1;            //体检人员的ID所在列
         public  int PNname_column=2;          //体检人员的名字所在列
         public  int pSex_column=3;            //性别所在列
         public  int pAge_column=4;            //年龄所在列
         public  int pItemSet_column=5;        //相关检查项的集合所在列
         public  int pItem_column=6;           //检查项所在列
         public  int pItemData_column=7;       //检查项的数据所在列
         public  int itemBegin_column = 5;     //检查项起始列




        //进度条更新函数
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

        private void openExcel(String filename)
        {
            app.Visible = false; 
            this.wbks = app.Workbooks;
            r_wbk= wbks.Add(filename);
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
        private void collectdata()
        {
            openExcel(txtboxPath.Text);
            //createExcel();
            //Sheets shs = w_wbk.Sheets;
            //_Worksheet _wsh = (_Worksheet)shs.get_Item(1);
            //_wsh.Name = "汇总";
            //_wsh.Cells[2, 2] = "测试";
            //CloseExcel(w_wbk);
         
            String curPeopleID;                                         //存储当前的检测用户
            String prePeopleID;                                         //存储前一个用户
            Sheets r_shs = r_wbk.Sheets;                                //获取excel工作簿的所有工作表
            _Worksheet r_wsh = (_Worksheet)r_shs.get_Item(1);           //获取第一个工作表
            long i=1;                                                               //记录行的位置
            if (Convert.ToString(((Range)r_wsh.Cells[i, pItemData_column]).Text) == "")               //第一行为空
                i = 2;
            int j=itemBegin_column;                                     //数据项的起始位置
            curPeopleID =Convert.ToString(((Range) r_wsh.Cells[i, pID_column]).Text);
            prePeopleID = "";
            peopledata curpeople = new peopledata();                    // 存取当前的体检人员的数据
            KeyValuePair<String,String> itemset_item;
            int oldrowindex;                                            //存在旧的index
            long allrow = r_wsh.UsedRange.CurrentRegion.Rows.Count;
            //while(i !=r_wsh.UsedRange.Cells.Rows.Count)  // 没读到行结束
            //((Range)r_wsh.Cells[i,pID_column]).Text.ToString() !=""
            this.label3.Text = "正在读取数据,请稍后……";
            while(i <= allrow)
            {
                if (curPeopleID == prePeopleID)//用户id没改变
                {
                     itemset_item  = new KeyValuePair<String,String>(Convert.ToString(((Range)r_wsh.Cells[i,pItemSet_column]).Text),Convert.ToString(((Range)r_wsh.Cells[i,pItem_column]).Text));
                     // this.label2.Text = i.ToString();
                     curpeople.element.Add(itemset_item, Convert.ToString(((Range)r_wsh.Cells[i, pItemData_column]).Text));     //增加检查项目
                     if (i == allrow)
                     {
                         peoples.Add(curpeople);
                     
                     }
                }
                else //当前用户不同与上一行的用户id
                {                        
                      //新用户添加基础信息            
                      oldrowindex = peoplesId.IndexOf(curPeopleID);

                      if (oldrowindex != -1)//已经记录过此用户
                      {
                          itemset_item = new KeyValuePair<String, String>(Convert.ToString(((Range)r_wsh.Cells[i, pItemSet_column]).Text), Convert.ToString(((Range)r_wsh.Cells[i, pItem_column]).Text));
                          peoples[oldrowindex].element.Add(itemset_item, Convert.ToString(((Range)r_wsh.Cells[i, pItemData_column]).Text));     //增加检查项目
                          oldrowindex = peoplesId.IndexOf(prePeopleID);
                          if (oldrowindex == -1)
                          {
                              peoplesId.Add(prePeopleID);
                              peoples.Add(curpeople);
                          }
                      }
                      else
                      {
                         // peoplesId.Add(curPeopleID);
                          if (prePeopleID != "")//不是第一个用户
                          {
                              oldrowindex = peoplesId.IndexOf(prePeopleID);
                              if (oldrowindex == -1)
                              {
                                  peoplesId.Add(prePeopleID);
                                  peoples.Add(curpeople);
                              }
                              curpeople = new peopledata();
                          }
                          //添加检查项目信息
                          itemset_item = new KeyValuePair<String, String>(Convert.ToString(((Range)r_wsh.Cells[i, pItemSet_column]).Text), Convert.ToString(((Range)r_wsh.Cells[i, pItem_column]).Text));
                          curpeople.Id = curPeopleID;
                          curpeople.name = Convert.ToString(((Range)r_wsh.Cells[i, PNname_column]).Text);
                          curpeople.sex = Convert.ToString(((Range)r_wsh.Cells[i, pSex_column]).Text);
                          curpeople.age = Convert.ToString(((Range)r_wsh.Cells[i, pAge_column]).Text);
                          curpeople.element.Add(itemset_item, Convert.ToString(((Range)r_wsh.Cells[i, pItemData_column]).Text));     //增加检查项目  
                          if (i == allrow)
                          {
                              peoples.Add(curpeople);

                          }
                      }
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
                if (i <= allrow)
                {
                    prePeopleID = curPeopleID;
                    curPeopleID = Convert.ToString(((Range)r_wsh.Cells[i, pID_column]).Text);
                }
            }
            app.AlertBeforeOverwriting = false;         // 屏蔽系统的alert
            r_wbk.Close();

        }


        //把数据写入一个工作簿
        private void writedata()
        {
            createExcel();
            Sheets w_shs = w_wbk.Sheets;                                               //获取excel工作簿的所有工作表
            _Worksheet w_wsh = (_Worksheet)w_shs.get_Item(1);                          //获取第一个工作表
            long i = 2;
            int lineheader = 1;//定义列头开始行
            int n = peoples.Count;
            this.label3.Text = "正在写入汇总结果，请稍后……";

            w_wsh.Cells[lineheader, pID_column] = "体检号";
            w_wsh.Cells[lineheader, PNname_column] = "姓名";
            w_wsh.Cells[lineheader, pSex_column] = "性别";
            w_wsh.Cells[lineheader, pAge_column] = "年龄";
            foreach (KeyValuePair<KeyValuePair<String, String>, int> kp in itemDic)
            {
                w_wsh.Cells[lineheader, kp.Value] = kp.Key.Value;       
            }

            foreach (peopledata p in peoples)
            { 
                w_wsh.Cells[i,pID_column]=p.Id;
                w_wsh.Cells[i,PNname_column] = p.name;
                w_wsh.Cells[i,pSex_column] = p.sex;
                w_wsh.Cells[i,pAge_column] = p.age;
                int itcolumn;
                foreach (KeyValuePair<KeyValuePair<String, String>, String> it in p.element)
                {
                    itcolumn = itemDic[it.Key];       //获取行号
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
               //this.Invoke(new InvokeDe)
                Thread fThread = new Thread(new ThreadStart(solve));        //开辟一个新的线
                fThread.IsBackground = true;
                fThread.Start();
                //fThread.Join();

                while (true)
                {
                    if (threadStart == true)
                    {
                        fThread.Abort();
                        threadStart = false;
                        peoples.Clear();
                        peoplesId.Clear();
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

                collectdata();      //读取数据
                writedata();        //写数据
                label3.Text = "完成";
                MessageBox.Show("处理完成");
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
                    pID_column =1;            //体检人员的ID所在列
                    PNname_column=2;          //体检人员的名字所在列
                    pSex_column=3;            //性别所在列
                    pAge_column=4;            //年龄所在列
                    pItemSet_column=5;        //相关检查项的集合所在列
                    pItem_column=6;           //检查项所在列
                    pItemData_column=7;       //检查项的数据所在列
                    itemBegin_column = 5;     //检查项起始列
            }
            else if (radioButton2.Checked == true)
            {
                pID_column = 1;            //体检人员的ID所在列
                PNname_column = 2;          //体检人员的名字所在列
                pSex_column = 3;            //性别所在列
                pAge_column = 4;            //年龄所在列
                pItemSet_column = 7;        //相关检查项的集合所在列
                pItem_column = 8;           //检查项所在列
                pItemData_column = 9;       //检查项的数据所在列
                itemBegin_column = 5;     //检查项起始列               
            }
            else
            {
                MessageBox.Show("请选择一种数据格式！");          
            }   


        }
    }
}
