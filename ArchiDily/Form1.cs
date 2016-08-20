using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ArchiDily
{
    public partial class MainForm : Form
    {
        private string DataFileLocation="";
        private DataSet[] DataSet_Target=new DataSet[0];//要打印的数据集合
        private PDF pdf = new PDF();

        public MainForm()
        {
            InitializeComponent();
        }

        private bool createTwoTabledPage()
        {
            return true;
        }

        private bool createThreeTabledPage()
        {
            return true;
        }

        /// <summary>
        /// 设置控件大小及位置
        /// </summary>
        private void reSizeControls(bool onLoad)
        {
            if (onLoad == true || this.panel_guidemode_step1.Visible == true)
            {
                this.panel_guidemode_step1.Size = this.toolStripContainer1.ContentPanel.Size;
                this.panel_guidemode_step1.Location = new Point(0, 0);

                this.button_chooseFileLocation.Location =
                    new Point(this.panel_guidemode_step1.Width / 2 - this.button_chooseFileLocation.Width / 2,
                    this.panel_guidemode_step1.Height / 2 - this.button_chooseFileLocation.Height);
                this.label_chooseFileLocaton.Location =
                    new Point(this.panel_guidemode_step1.Width / 2 - (this.label_chooseFileLocaton.Width + this.label_fileLocaton.Width) / 2,
                    this.button_chooseFileLocation.Location.Y + this.button_chooseFileLocation.Height + 10);
                this.label_fileLocaton.Location =
                    new Point(this.label_chooseFileLocaton.Location.X + this.label_chooseFileLocaton.Width,
                    this.label_chooseFileLocaton.Location.Y);
                this.button_nextStep1.Location =
                    new Point(this.panel_guidemode_step1.Width / 2 - this.button_nextStep1.Width / 2,
                    this.label_chooseFileLocaton.Location.Y + this.label_chooseFileLocaton.Height + 10);
            }

            if (onLoad == true || this.panel_guidemode_step2.Visible == true)
            {
                this.panel_guidemode_step2.Size = this.toolStripContainer1.ContentPanel.Size;
                this.panel_guidemode_step2.Location = new Point(0, 0);

                this.listBox_dataTable.Width = this.panel_guidemode_step2.Width / 6;

                this.listBox_dataTable.Location = new Point(this.panel_guidemode_step2.Width / 2 - this.listBox_dataTable.Width / 2,
                    this.panel_guidemode_step2.Height / 2 - this.listBox_dataTable.Height / 2);
                this.label_chooseDataTable.Location = new Point(this.listBox_dataTable.Location.X-this.label_chooseDataTable.Width-3,
                    this.listBox_dataTable.Location.Y);
                this.button_nextStep2.Location = new Point(this.panel_guidemode_step2.Width / 2 - this.button_nextStep2.Width / 2,
                    this.listBox_dataTable.Location.Y + this.listBox_dataTable.Height + 5);
                
            }

            if (onLoad == true || this.panel_guidemode_step3.Visible == true)
            {
                this.panel_guidemode_step3.Size = this.toolStripContainer1.ContentPanel.Size;
                this.panel_guidemode_step3.Location = new Point(0, 0);
            }

            if (onLoad == true || this.panel_guidemode_step4.Visible == true)
            {
                this.panel_guidemode_step4.Size = this.toolStripContainer1.ContentPanel.Size;
                this.panel_guidemode_step4.Location = new Point(0, 0);
            }

        }

        private bool selectTable()
        {
            if (DataFileLocation == "") return false;
            string fileType = System.IO.Path.GetExtension(DataFileLocation);
            string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                DataFileLocation + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            string sql_F = "Select * FROM [{0}]";

            OleDbConnection conn = null;
            OleDbDataAdapter da = null;
            DataTable dtSheetName = null;


            try
            {
                conn = new OleDbConnection(connStr);
                conn.Open();
            }
            catch (Exception)
            {
                connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                    DataFileLocation + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
                conn = new OleDbConnection(connStr);
                conn.Open();
            }

            try
            {
                dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                da = new OleDbDataAdapter();
                this.listBox_dataTable.Items.Clear();
                for (int i = 0; i < dtSheetName.Rows.Count; i++)
                {
                    this.listBox_dataTable.Items.Add((string)dtSheetName.Rows[i]["TABLE_NAME"]);
                }
                this.listBox_dataTable.Height = (this.listBox_dataTable.Items.Count + 1) * this.listBox_dataTable.ItemHeight;
                reSizeControls(true);

                DataSet_Target = new DataSet[dtSheetName.Rows.Count];
                for (int i = 0; i < dtSheetName.Rows.Count; i++)
                {
                    //SheetName = (string)dtSheetName.Rows[i]["TABLE_NAME"];
                    //MessageBox.Show(SheetName);

                    //if (SheetName.Contains("$") && !SheetName.Replace("'", "").EndsWith("$"))
                   // {
                    //    continue;
                    //}

                    da.SelectCommand = new OleDbCommand(String.Format(sql_F, (string)dtSheetName.Rows[i]["TABLE_NAME"]), conn);
                    DataSet dsItem = new DataSet();
                    da.Fill(dsItem, (string)dtSheetName.Rows[i]["TABLE_NAME"]);

                    DataSet_Target[i] = new DataSet();
                    DataSet_Target[i].Tables.Add(dsItem.Tables[0].Copy());
                }
            }
            //catch (Exception ex) { this.toolStripStatusLabel1.Text = ex.Message; return false; }
            finally
            {
                // 关闭连接
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    da.Dispose();
                    conn.Dispose();
                }
            }
            return true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            /**/
            this.menuStrip1.Visible = false;
            this.toolStripStatusLabel1.Text = "";
            this.label_fileLocaton.Visible = this.label_chooseFileLocaton.Visible = this.button_nextStep1.Visible = false;
            this.panel_guidemode_step1.Visible = false;
            this.panel_guidemode_step2.Visible = false;
            this.panel_guidemode_step3.Visible = false;
            this.panel_guidemode_step4.Visible = false;

            /*设置控件尺寸及位置
            this.panel_guidemode_step1.Size = this.toolStripContainer1.ContentPanel.Size;
            this.panel_guidemode_step1.Location = new Point(0, 0);*/
            reSizeControls(true);
        }

        private void toolStripButton2_MouseEnter(object sender, EventArgs e)
        {
            this.toolStripStatusLabel1.Text = "启动向导模式";
        }

        private void toolStripButton2_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatusLabel1.Text = "";
        }

        private void toolStripButton1_MouseEnter(object sender, EventArgs e)
        {
            this.toolStripStatusLabel1.Text = "打开数据文件";
        }

        private void toolStripButton1_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatusLabel1.Text = "";
        }

        private void panel_guidemode_step1_Resize(object sender, EventArgs e)
        {
            /*
            this.button_chooseFileLocation.Location =
                new Point(this.panel_guidemode_step1.Width / 2 - this.button_chooseFileLocation.Width / 2,
                this.panel_guidemode_step1.Height / 2 - this.button_chooseFileLocation.Height);
            this.label_chooseFileLocaton.Location =
                new Point(this.panel_guidemode_step1.Width / 2 - (this.label_chooseFileLocaton.Width + this.label_fileLocaton.Width) / 2,
                this.button_chooseFileLocation.Location.Y + this.button_chooseFileLocation.Height + 10);
            this.label_fileLocaton.Location =
                new Point(this.label_chooseFileLocaton.Location.X + this.label_chooseFileLocaton.Width,
                this.label_chooseFileLocaton.Location.Y);
            this.button_nextStep1.Location =
                new Point(this.panel_guidemode_step1.Width / 2 - this.button_nextStep1.Width / 2,
                this.label_chooseFileLocaton.Location.Y + this.label_chooseFileLocaton.Height + 10);
                */
            reSizeControls(false);
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            /*设置控件尺寸及位置
            this.panel_guidemode_step1.Size = this.toolStripContainer1.ContentPanel.Size;
            this.panel_guidemode_step1.Location = new Point(0, 0);*/

            reSizeControls(false);
        }

        private void button_chooseFileLocation_Click(object sender, EventArgs e)
        {
            this.toolStripButton_openFile.PerformClick();
            /*
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel 文件(*.xls, *.xlsx)|*.xls; *.xlsx|Excel 97 - 2003文件(*.xls)|*.xls|Excel 2007 - 2016文件(*.xlsx)|*.xlsx|所有文件|*.*";
            ofd.ShowDialog(this);
            if (ofd.FileName == "")
            {
                this.toolStripStatusLabel1.Text = "未选择数据文件";
            }
            else
            {
                //设定数据文件地址
                this.DataFileLocation = ofd.FileName;
                this.label_fileLocaton.Text = DataFileLocation;
                //相关控件设定
                this.label_fileLocaton.Visible = this.label_chooseFileLocaton.Visible = this.button_nextStep1.Visible = true;
                this.label_chooseFileLocaton.Location =
                new Point(this.panel_guidemode_step1.Width / 2 - (this.label_chooseFileLocaton.Width + this.label_fileLocaton.Width) / 2,
                this.button_chooseFileLocation.Location.Y + this.button_chooseFileLocation.Height + 10);
                this.label_fileLocaton.Location =
                    new Point(this.label_chooseFileLocaton.Location.X + this.label_chooseFileLocaton.Width,
                    this.label_chooseFileLocaton.Location.Y);
            }
            */

        }

        private void toolStripButton_exit_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void toolStripButton_GuideMode_Click(object sender, EventArgs e)
        {
            this.panel_guidemode_step1.Visible = true;
            this.panel_guidemode_step2.Visible = false;
            this.panel_guidemode_step3.Visible = false;
            this.panel_guidemode_step4.Visible = false;
        }

        private void button_nextStep1_Click(object sender, EventArgs e)
        {
            if (selectTable())
            {
                this.panel_guidemode_step1.Visible = false;
                this.panel_guidemode_step2.Visible = true;
                this.panel_guidemode_step3.Visible = false;
                this.panel_guidemode_step4.Visible = false;

                this.listBox_dataTable.SetSelected(0, true);
            }

            
        }
        
        

        private void toolStripButton_openFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel 文件(*.xls, *.xlsx)|*.xls; *.xlsx|Excel 97 - 2003文件(*.xls)|*.xls|Excel 2007 - 2016文件(*.xlsx)|*.xlsx|所有文件|*.*";
            ofd.ShowDialog(this);
            if (ofd.FileName == "")
            {
                this.toolStripStatusLabel1.Text = "未选择数据文件";
            }
            else
            {
                //设定数据文件地址
                this.DataFileLocation = ofd.FileName;
                this.label_fileLocaton.Text = DataFileLocation;
                //相关控件设定
                this.label_fileLocaton.Visible = this.label_chooseFileLocaton.Visible = this.button_nextStep1.Visible = true;
                this.label_chooseFileLocaton.Location =
                new Point(this.panel_guidemode_step1.Width / 2 - (this.label_chooseFileLocaton.Width + this.label_fileLocaton.Width) / 2,
                this.button_chooseFileLocation.Location.Y + this.button_chooseFileLocation.Height + 10);
                this.label_fileLocaton.Location =
                    new Point(this.label_chooseFileLocaton.Location.X + this.label_chooseFileLocaton.Width,
                    this.label_chooseFileLocaton.Location.Y);
            }

            this.panel_guidemode_step1.Visible = true;
            this.panel_guidemode_step2.Visible = false;
            this.panel_guidemode_step3.Visible = false;
            this.panel_guidemode_step4.Visible = false;
        }

        private void button_nextStep2_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(this.listBox_dataTable.SelectedItem.ToString());
            
            if (this.listBox_dataTable.SelectedIndex.ToString()=="-1")
            {
                this.listBox_dataTable.SelectedIndex = 0;
            }
            this.panel_guidemode_step1.Visible = false;
            this.panel_guidemode_step2.Visible = false;
            this.panel_guidemode_step3.Visible = true;
            this.panel_guidemode_step4.Visible = false;

            //this.dataGridView1.DataSource = this.DataSet_Target[this.listBox_dataTable.SelectedIndex].Tables[0];
        }

        private void button_nextStep3_Click(object sender, EventArgs e)
        {
            this.panel_guidemode_step1.Visible = false;
            this.panel_guidemode_step2.Visible = false;
            this.panel_guidemode_step3.Visible = false;
            this.panel_guidemode_step4.Visible = true;
        }

        private void toolStripButton_chooseDataTable_Click(object sender, EventArgs e)
        {
            if (selectTable())
            {
                this.panel_guidemode_step1.Visible = false;
                this.panel_guidemode_step2.Visible = true;
                this.panel_guidemode_step3.Visible = false;
                this.panel_guidemode_step4.Visible = false;

                this.listBox_dataTable.SetSelected(0, true);                
            }
        }

        private void button_done_Click(object sender, EventArgs e)
        {
            
        }

        private void radioButton_DOUBLE_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_DOUBLE.Checked)
            {
                pdf.blocks = 2;
            }
        }

        private void radioButton_TRIPLE_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton_TRIPLE.Checked)
            {
                pdf.blocks = 3;
            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            this.label5.Text = "即：" + this.textBox1.Text + "档字第：X号";
        }
    }
}
