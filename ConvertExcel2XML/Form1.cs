using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Xml;

namespace ConvertExcel2XML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string filename = "";
        string xlsConnStr;
        string xmlFilename = "";


        private void btnBrowse_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Microsoft Excel|*.xls;*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filename = openFileDialog1.FileName;
                txtSource.Text = openFileDialog1.FileName;
                FileInfo f = new FileInfo(filename);
                string ext = f.Extension;
                if (ext.Equals(".xls"))
                {
                    xlsConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties=Excel 8.0";
                }
                if (ext.Equals(".xlsx"))
                {
                    xlsConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties=Excel 12.0";
                }
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            if (txtSource.Text.Trim() == "")
            {
                MessageBox.Show("Pls chose source!!!");
                return;
            }
            try
            {
                saveFileDialog1.Filter = "XML|*.xml";
                if (DialogResult.OK != saveFileDialog1.ShowDialog()) { return; }
                xmlFilename = saveFileDialog1.FileName;

                XmlTextWriter writer = new XmlTextWriter(xmlFilename, null);
               
                OleDbConnection xlsConn = new OleDbConnection(xlsConnStr);
                xlsConn.Open();

                //lay tat ca cac ten sheet
                DataTable dtSheet = xlsConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                DataSet ds = new DataSet();
                foreach (DataRow drSheet in dtSheet.Rows)
                {
                    if (drSheet["TABLE_NAME"].ToString().Contains("$"))
                    {
                        string sheetname = drSheet["TABLE_NAME"].ToString();

                        OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [" + sheetname + "]", xlsConn);
                        
                        da.Fill(ds, sheetname.Remove(sheetname.Length - 1));
                    }
                }
                ds.WriteXml(writer);
                writer.Close();
                MessageBox.Show("Đã chuyển đổi thành công!!!");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            Form2 obj = new Form2(xmlFilename);
            obj.ShowDialog();
        }







    }
}
