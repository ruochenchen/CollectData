using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace checkdataCollect
{
    class peopledata
    {
        public String name;
        public String sex;
        public String age;
        public Dictionary<KeyValuePair<String,String>, String> element = new Dictionary<KeyValuePair<String,String>,String>();

    }
    class peopleSet
    {
        public Dictionary<string, peopledata> peopledic;

        public peopleSet(int style)
        {
            peopledic = new Dictionary<string, peopledata>();
            ExcelStyleChange(style);
        }


        public void ExcelStyleChange(int style)
        {
            if (style == 1)
            {
                pID_column = 1;            //体检人员的ID所在列
                PNname_column = 2;          //体检人员的名字所在列
                pSex_column = 3;            //性别所在列
                pAge_column = 4;            //年龄所在列
                pItemSet_column = 5;        //相关检查项的集合所在列
                pItem_column = 6;           //检查项所在列
                pItemData_column = 7;       //检查项的数据所在列

                //写入数据定义
                itemBegin_column = 5;     //检查项起始列
            }
            else
            {
                pID_column = 1;            //体检人员的ID所在列
                PNname_column = 2;          //体检人员的名字所在列
                pSex_column = 4;            //性别所在列
                pAge_column = 5;            //年龄所在列
                pItemSet_column = 7;        //相关检查项的集合所在列
                pItem_column = 9;           //检查项所在列
                pItemData_column = 11;       //检查项的数据所在列

                //写入数据定义
                itemBegin_column = 5;     //检查项起始列
            }
        }

        //excel里面一些位置的定义
        public int pID_column = 1;            //体检人员的ID所在列
        public int PNname_column = 2;          //体检人员的名字所在列
        public int pSex_column = 3;            //性别所在列
        public int pAge_column = 4;            //年龄所在列
        public int pItemSet_column = 5;        //相关检查项的集合所在列
        public int pItem_column = 6;           //检查项所在列
        public int pItemData_column = 7;       //检查项的数据所在列


        //写入数据定义
        public int itemBegin_column = 5;     //检查项起始列
    }
}
