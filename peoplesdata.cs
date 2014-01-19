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
        public Dictionary<string, peopledata> peopledic = new   Dictionary<string,peopledata>();

    }
}
