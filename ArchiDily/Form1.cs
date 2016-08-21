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

        private bool setStep3()
        {
            if (DataSet_Target != null)
            {
                for (int time = 1; time <= 8; time++)
                {
                    DataTable ADt;
                    DataColumn ADC1;
                    DataColumn ADC2;
                    ADt = new DataTable();
                    ADC1 = new DataColumn("F_ID", typeof(int));
                    ADC2 = new DataColumn("F_Name", typeof(string));
                    ADt.Columns.Add(ADC1);
                    ADt.Columns.Add(ADC2);
                    for (int i = 0; i < DataSet_Target[listBox_dataTable.SelectedIndex].Tables[0].Columns.Count; i++)
                    {
                        DataRow ADR = ADt.NewRow();
                        ADR[0] = i + 1;
                        ADR[1] = DataSet_Target[listBox_dataTable.SelectedIndex].Tables[0].Columns[i].ToString();
                        ADt.Rows.Add(ADR);
                    }
                    
                    switch (time)
                    {
                        case 1://档案接收单位
                            comboBox_forwarding.DisplayMember = "F_Name";
                            comboBox_forwarding.ValueMember = "F_ID";
                            comboBox_forwarding.DataSource = ADt;
                            break;
                        case 2://档案号
                            comboBox_filenum.DisplayMember = "F_Name";
                            comboBox_filenum.ValueMember = "F_ID";
                            comboBox_filenum.DataSource = ADt;
                            break;
                        case 3://var1
                            comboBoxVar1.DisplayMember = "F_Name";
                            comboBoxVar1.ValueMember = "F_ID";
                            comboBoxVar1.DataSource = ADt;
                            break;
                        case 4://var2
                            comboBoxVar2.DisplayMember = "F_Name";
                            comboBoxVar2.ValueMember = "F_ID";
                            comboBoxVar2.DataSource = ADt;
                            break;
                        case 5://var3
                            comboBoxVar3.DisplayMember = "F_Name";
                            comboBoxVar3.ValueMember = "F_ID";
                            comboBoxVar3.DataSource = ADt;
                            break;
                        case 6://var4
                            comboBoxVar4.DisplayMember = "F_Name";
                            comboBoxVar4.ValueMember = "F_ID";
                            comboBoxVar4.DataSource = ADt;
                            break;
                        case 7://var5
                            comboBoxVar5.DisplayMember = "F_Name";
                            comboBoxVar5.ValueMember = "F_ID";
                            comboBoxVar5.DataSource = ADt;
                            break;
                        case 8://var6
                            comboBoxVar6.DisplayMember = "F_Name";
                            comboBoxVar6.ValueMember = "F_ID";
                            comboBoxVar6.DataSource = ADt;
                            break;
                        default: return false;
                    }
                }
                return true;
            }
            else { return false; }
        }
        private void button_nextStep2_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(this.listBox_dataTable.SelectedItem.ToString());
            if (setStep3())
            {
                if (this.listBox_dataTable.SelectedIndex.ToString() == "-1")
                {
                    this.listBox_dataTable.SelectedIndex = 0;
                }
                this.panel_guidemode_step1.Visible = false;
                this.panel_guidemode_step2.Visible = false;
                this.panel_guidemode_step3.Visible = true;
                this.panel_guidemode_step4.Visible = false;

                //this.dataGridView1.DataSource = this.DataSet_Target[this.listBox_dataTable.SelectedIndex].Tables[0];
            }
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
            //filePath
            //pdf.filepath = "test.pdf";
            if (radioButtonFileNameTime.Checked)
            {
                pdf.filepath = DateTime.Now.Year.ToString("D4") + DateTime.Now.Month.ToString("D2") + DateTime.Now.Day.ToString("D2") +
                  DateTime.Now.Hour.ToString("D2") + DateTime.Now.Minute.ToString("D2") + DateTime.Now.Second.ToString("D2")+".pdf";
            }
            if(radioButtonFileNameFixed.Checked)
            {
                pdf.filepath = textBoxFileNameFixed.Text + ".pdf";
            }
            //title
            pdf.title = this.textBox_Title.Text;
            //forwarding
            if (radioButton_forwarding_binding.Checked) { pdf.forwarding = comboBox_forwarding.SelectedIndex.ToString(); }
            if (radioButton_forwarding_typein.Checked) { pdf.forwarding = textBox_Forwarding.Text; }
            if (radioButton_forwarding_hide.Checked) { pdf.forwarding = ""; }
            //filenum
            if (radioButton_filenum_auto.Checked) { pdf.filenum = comboBox_filenum.SelectedIndex.ToString(); }
            if (radioButton_filenum_code.Checked) { pdf.filenum = textBox_filenum.Text; }
            if (radioButton_filenum_hide.Checked) { pdf.filenum = ""; }
            //filetype
            pdf.filetype = textBoxFileType.Text;
            //maintext1
            pdf.mainText1 = this.textBox_MainText1.Text;
            //maintext2
            pdf.mainText2 = this.textBox_MainText2.Text;
            //maintext3
            pdf.mainText3 = this.textBox_MainText3.Text;
            //maintext4
            pdf.mainText4 = this.textBox_MainText4.Text;
            //maintext5
            pdf.mainText5 = this.textBox_MainText5.Text;
            //name
            //peoplecount
            pdf.peopleCount = "1";
            //filecount1
            pdf.fileCount1 = "1";
            //fileconunt2
            pdf.fileCount2 = "1";
            //sending
            pdf.sending = this.textBox_Sending.Text;
            //date
            pdf.date = DateTime.Now.Year.ToString() + "年" + DateTime.Now.Month.ToString() + "月" + DateTime.Now.Day.ToString() + "日";
            //details
            //detailWidth
                        List<float> widthTemp = new List<float>();
            if (checkBoxVar1.Checked) { widthTemp.Add(float.Parse(numericUpDownVar1.Value.ToString())); }
            if (checkBoxVar2.Checked) { widthTemp.Add(float.Parse(numericUpDownVar2.Value.ToString())); }
            if (checkBoxVar3.Checked) { widthTemp.Add(float.Parse(numericUpDownVar3.Value.ToString())); }
            if (checkBoxVar4.Checked) { widthTemp.Add(float.Parse(numericUpDownVar4.Value.ToString())); }
            if (checkBoxVar5.Checked) { widthTemp.Add(float.Parse(numericUpDownVar5.Value.ToString())); }
            if (checkBoxVar6.Checked) { widthTemp.Add(float.Parse(numericUpDownVar6.Value.ToString())); }
            pdf.setTableWidth = widthTemp.ToArray();
            
            //string[,] s = new string[,] { };
            pdf.details = new List<List<string>>();
            List<string> tmpDetail = new List<string>();
            int i_tmp_1 = 0,printedpagenum=DataSet_Target[this.listBox_dataTable.SelectedIndex].Tables[0].Rows.Count;
            //if (checkBoxVar1.Checked) { pdf.details_0.SetValue(textBox_Var1.Text, 0, 0); i_tmp_1++; }
            //if (checkBoxVar2.Checked) { pdf.details_0.SetValue(textBox_Var2.Text, 0, i_tmp_1); i_tmp_1++; }
            //if (checkBoxVar3.Checked) { pdf.details_0.SetValue(textBox_Var3.Text, 0, i_tmp_1); i_tmp_1++; }
            //if (checkBoxVar4.Checked) { pdf.details_0.SetValue(textBox_Var4.Text, 0, i_tmp_1); i_tmp_1++; }
            //if (checkBoxVar5.Checked) { pdf.details_0.SetValue(textBox_Var5.Text, 0, i_tmp_1); i_tmp_1++; }
            //if (checkBoxVar6.Checked) { pdf.details_0.SetValue(textBox_Var6.Text, 0, i_tmp_1); }
            if (checkBoxVar1.Checked) { tmpDetail.Add(textBox_Var1.Text); i_tmp_1++; }
            if (checkBoxVar2.Checked) { tmpDetail.Add(textBox_Var2.Text); i_tmp_1++; }
            if (checkBoxVar3.Checked) { tmpDetail.Add(textBox_Var3.Text); i_tmp_1++; }
            if (checkBoxVar4.Checked) { tmpDetail.Add(textBox_Var4.Text); i_tmp_1++; }
            if (checkBoxVar5.Checked) { tmpDetail.Add(textBox_Var5.Text); i_tmp_1++; }
            if (checkBoxVar6.Checked) { tmpDetail.Add(textBox_Var6.Text); }
            pdf.details.Add(tmpDetail);

            for (int i_tmp_2 = 0; i_tmp_2 < printedpagenum; i_tmp_2++)
            {
                tmpDetail = new List<string>();
                //for (int i_tmp_3 = 0; i_tmp_3 <= i_tmp_1; i_tmp_3++)
                //{
                    //pdf.details_0.SetValue("test", i_tmp_2, i_tmp_3);
                  //  tmpDetail.Add("testvalue" + i_tmp_3.ToString());
                //}
                if (radioButtonRangeAll.Checked)
                {
                    if (checkBoxVar1.Checked) { tmpDetail.Add(DataSet_Target[this.listBox_dataTable.SelectedIndex].Tables[0].Rows[i_tmp_2][comboBoxVar1.SelectedIndex].ToString()); }
                    if (checkBoxVar2.Checked) { tmpDetail.Add(DataSet_Target[this.listBox_dataTable.SelectedIndex].Tables[0].Rows[i_tmp_2][comboBoxVar2.SelectedIndex].ToString()); }
                    if (checkBoxVar3.Checked) { tmpDetail.Add(DataSet_Target[this.listBox_dataTable.SelectedIndex].Tables[0].Rows[i_tmp_2][comboBoxVar3.SelectedIndex].ToString()); }
                    if (checkBoxVar4.Checked) { tmpDetail.Add(DataSet_Target[this.listBox_dataTable.SelectedIndex].Tables[0].Rows[i_tmp_2][comboBoxVar4.SelectedIndex].ToString()); }
                    if (checkBoxVar5.Checked) { tmpDetail.Add(DataSet_Target[this.listBox_dataTable.SelectedIndex].Tables[0].Rows[i_tmp_2][comboBoxVar5.SelectedIndex].ToString()); }
                    if (checkBoxVar6.Checked) { tmpDetail.Add(DataSet_Target[this.listBox_dataTable.SelectedIndex].Tables[0].Rows[i_tmp_2][comboBoxVar6.SelectedIndex].ToString()); }
                }
                if (radioButtonRangeSelected.Checked)
                { }
                if (radioButtonRangeIdentical.Checked)
                { }
                
                pdf.details.Add(tmpDetail);
            }
            //receipt
            pdf.receipt = this.textBox_Receipt.Text;
            //RForwarding
            pdf.receiptForward = this.textBox_RForward.Text;
            //rtext1
            pdf.receiptText1 = this.textBox_RText1.Text;
            //rtext2
            pdf.receiptText2 = this.textBox_RText2.Text;
            //rtext3
            pdf.receiptText3 = this.textBox_RText3.Text;
            //rtext4
            pdf.receiptText4 = this.textBox_RText4.Text;
            //rtext5
            pdf.receiptText5 = this.textBox_RText5.Text;
            //rtap
            pdf.receiptTap = this.textBox_RTap.Text;
            //rsign
            pdf.receiptSign = this.textBox_RSign.Text;
            //rdate
            pdf.receiptDate = this.textBox_RDate.Text;
            //address
            pdf.address = this.textBox_Address.Text;
            //zipcode
            pdf.zipcode = this.textBox_ZipCode.Text;



            if (pdf.CreatePDF())
            {
                MessageBox.Show("success");
            }
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
                this.groupBox2.Size = new Size(717, 419);
                this.groupBox2.Location = new Point(11, 90);
                this.groupBox2.Visible = true;
                this.groupBox3.Visible = false;
            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            this.label5.Text = "即：" + this.textBoxFileType.Text + "档字第：X号";
        }

        private void radioButton_forwarding_binding_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_forwarding_binding.Checked)
            {
                pdf.forwardingType = 1;
            }
        }

        private void radioButton_forwarding_typein_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_forwarding_typein.Checked)
            {
                pdf.forwardingType = 2;
            }
        }

        private void radioButton_forwarding_hide_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_forwarding_hide.Checked)
            {
                pdf.forwardingType = 3;
            }
        }

        private void radioButton_filenum_auto_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_filenum_auto.Checked)
            {
                pdf.SetFileNum(new string[] { "1", "1" });
            }
        }

        private void radioButton_filenum_code_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_filenum_code.Checked)
            {
                if (checkBox_filenum_fixed.Checked)
                {
                    pdf.SetFileNum(new string[] { "2", this.numericUpDown_filenum_fixed.Value.ToString() });
                }
                else
                {
                    pdf.SetFileNum(new string[] { "2", "1" });
                }
            }
        }

        private void radioButton_filenum_hide_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton_filenum_hide.Checked)
            {
                pdf.SetFileNum(new string[] { "3", "1" });
            }
        }

        private void radioButtonVar1Bind_CheckedChanged(object sender, EventArgs e)
        {
            //
        }

        private void comboBox_forwarding_SelectedIndexChanged(object sender, EventArgs e)
        {
            radioButton_forwarding_binding.Checked = true;
        }

        private void textBox_Forwarding_TextChanged(object sender, EventArgs e)
        {
            radioButton_forwarding_typein.Checked = true;
        }

        private void comboBox_filenum_SelectedIndexChanged(object sender, EventArgs e)
        {
            radioButton_filenum_auto.Checked = true;
        }

        private void textBox_filenum_TextChanged(object sender, EventArgs e)
        {
            radioButton_filenum_code.Checked = true;
        }
    }
}
