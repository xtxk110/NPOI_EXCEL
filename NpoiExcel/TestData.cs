using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace NpoiExcel
{
    class TestData
    {
        [Description("姓名")]
        public string Name { get; set; }
        [Description("年龄")]
        public int Age { get; set; }
        [Description("生日")]
        public DateTime Birthday {get;set;}
    }
    class ComplexData
    {
        public string Name { get; set; }
        public bool flag { get; set; }
        public int Age { get; set; }
        public DateTime Birthday { get; set; }
        public Detail Detail { get; set; }
    }
    class Detail
    {
        public Address HomeAddr { get; set; }
        public string Tel { get; set; }
    }
    class Address
    {
        public string CommonAddr { get; set; }
        public string PostNo { get; set; }
    }

}
