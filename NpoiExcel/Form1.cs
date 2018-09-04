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
        List<ComplexData> list1 = new List<ComplexData>();
        public Form1()
        {
            InitializeComponent();
            list.Add(new TestData { Name = "zhangshan", Age = 40, Birthday = DateTime.Now });
            list.Add(new TestData { Name = "lisi", Age = 28, Birthday = DateTime.Now });
            list.Add(new TestData { Name = "wangwu", Age = 36, Birthday = DateTime.Now });
            list1.Add(new ComplexData { Name = "zhangshan", Age = 40, Birthday = DateTime.Now, Detail = new Detail { Tel = "02885394654", HomeAddr = new Address { CommonAddr = "黄河村", PostNo = "610048" } } });
            DataTable dt = new DataTable("TestTable");
            dt.Columns.Add("Name");
            dt.Columns.Add("Age");
            dt.Columns.Add("Birthday");
            DataRow dr = dt.NewRow();
            dr[0] = "zhangshan";
            dr[1] = 50;
            dr[2] = DateTime.Now;
            dt.Rows.Add(dr);

            NpoiHelper.GetInstance().ExportToExcel(dt);
            //NpoiHelper.GetInstance().ExportToExcel<TestData>(list);
            //NpoiHelper.GetInstance().ExportToExcel<ComplexData>(list1);

            //string file = AppDomain.CurrentDomain.BaseDirectory + "excel.xls";
            //NpoiHelper.GetInstance().ImportFromExcel(file,"0",0);
        }
    }
}
