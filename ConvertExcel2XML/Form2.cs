using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ConvertExcel2XML
{
    public partial class Form2 : Form
    {
        string path = "";
        DataSet ds;
        public Form2()
        {
            InitializeComponent();
        }

        public Form2(string path):this()
        {
            this.path = path;
        }

        private void cbTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            dgv.DataSource = ds.Tables[cbTable.Text];
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            try
            {
                ds = new DataSet();
                ds.ReadXml(path);
                int count = ds.Tables.Count;
                for (int i = 0; i < count; i++)
                {
                    cbTable.Items.Add(ds.Tables[i].TableName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                
            }
            

        }
    }
}
