using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NpoiExcel
{
    public partial class Form1 : Form
    {
        List<TestData> list = new List<TestData>();
        public Form1()
        {
            InitializeComponent();
            list.Add(new TestData { Name = "zhangshan", Age = 40, Birthday = DateTime.Now });
            list.Add(new TestData { Name = "lisi", Age = 28, Birthday = DateTime.Now });
            list.Add(new TestData { Name = "wangwu", Age = 36, Birthday = DateTime.Now });

            NpoiHelper.GetInstance().ExportToExcel<TestData>(list);
        }
    }
}
