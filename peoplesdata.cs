using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace checkdataCollect
{
    class peopledata
    {
        public String Id;
        public String name;
        public String sex;
        public String age;
        public Dictionary<KeyValuePair<String,String>, String> element = new Dictionary<KeyValuePair<String,String>,String>();

        //重写Equals，根据对象类型重载Equals
        public override bool Equals(object obj)
        {
            // If parameter cannot be cast to ThreeDPoint return false:
            if (obj is peopledata) return Equals((peopledata)obj);
            return base.Equals(obj);

        }
        //重写Equals，用id来判定数据是否重复
        public bool Equals(peopledata other)
        {
            return base.Equals(other.Id);
        }

        //重写get'Hashcode
        public override int GetHashCode()
        {
            return this.Id.GetHashCode();
        }


        //重载==操作符
        //public static bool operator ==(peopledata c1, peopledata c2)
        //{
        //    return c1.Equals(c2);
        //}

    }
}
