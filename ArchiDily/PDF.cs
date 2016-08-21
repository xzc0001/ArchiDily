using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace ArchiDily
{
    public class PDF
    {
        #region 属性变量定义
        private string FilePath;//文件名
        private string Title;//页面标题
        private string Forwarding;//接收单位
        private string FileNum, FileNum_Edited;//档案号
        private string FileType;//档案类型
        private string MainText1;//兹转去
        private string Name;//姓名
        private string MainText2;//同志等 
        private string PeopleCount;//人数
        private string MainText3;// 人档案材料 
        private string FileCount1;//袋数
        private string MainText4;//袋
        private string FileCount2;//件数
        private string MainText5;//份，请查收并将回执退回为盼。
        private string Sending;//落款
        private string Date;//日期
        private string[,] Details_0;//下方表格详情
        private List<List<string>> Details;
        private string Receipt;//回执
        private string ReceiptForward;//回执接收单位
        private string ReceiptText1;//收到你校
        private string ReceiptText2;//转来的
        private string ReceiptText3;//同志的档案材料共计
        private string ReceiptText4;//份已查对无误。
        private string ReceiptText5;//此复
        private string ReceiptTap;//收件机关盖章
        private string ReceiptSign;//签字
        private string ReceiptDate;//年月日
        private string Address;//表后地址
        private string Zipcode;//邮编

        private int PageCount_temp = 1;//
        private int I_column;//
        private int Blocks;
        private int ForwardingType;
        private int FileNumType;
        private int FileNumLength;
        private float TitlePaddingTop = 0f;
        private float[] DetailTableWidth;
        /// <summary>
        /// 文件
        /// </summary>
        public string filepath
        {
            get { return FilePath; }
            set { FilePath = value; }
        }

        public string title
        {
            get { return Title; }
            set { Title = value; }
        }

        public string forwarding
        {
            get { return Forwarding; }
            set { Forwarding = value; }
        }

        public string filenum
        {
            get { return FileNum; }
            set
            {
                FileNum = value;
                //FileNum_Edited = FileType + "档字：" + FileNum + "号";
                //FileNum_Edited=Select from Database or count;
            }
        }

        public string filetype
        {
            get { return FileType; }
            set { FileType = value; }
        }

        public string mainText1
        {
            get { return MainText1; }
            set { MainText1 = value; }
        }

        public string name
        {
            get { return Name; }
            set { Name = value; }
        }

        public string mainText2
        {
            get { return MainText2; }
            set { MainText2 = value; }
        }

        public string peopleCount
        {
            get { return PeopleCount; }
            set { PeopleCount = value; }
        }

        public string mainText3
        {
            get { return MainText3; }
            set { MainText3 = value; }
        }

        public string fileCount1
        {
            get { return FileCount1; }
            set { FileCount1 = value; }
        }

        public string mainText4
        {
            get { return MainText4; }
            set { MainText4 = value; }
        }

        public string fileCount2
        {
            get { return FileCount2; }
            set { FileCount2 = value; }
        }

        public string mainText5
        {
            get { return MainText5; }
            set { MainText5 = value; }
        }

        public string sending
        {
            get { return Sending; }
            set { Sending = value; }
        }

        public string date
        {
            get { return Date; }
            set { Date = value; }
        }

        public string[,] details_0
        {
            get { return Details_0; }
            set
            {
                Details_0 = value;
                this.I_column = value.GetLength(1);
            }
        }

        public List<List<string>> details
        {
            get { return Details; }
            set
            {
                Details = value;
                this.I_column = value.Count;
            }
        }

        public string receipt
        {
            get { return Receipt; }
            set { Receipt = value; }
        }

        public string receiptForward
        {
            get { return ReceiptForward; }
            set { ReceiptForward = value; }
        }

        public string receiptText1
        {
            get { return ReceiptText1; }
            set { ReceiptText1 = value; }
        }

        public string receiptText2
        {
            get { return ReceiptText2; }
            set { ReceiptText2 = value; }
        }

        public string receiptText3
        {
            get { return ReceiptText3; }
            set { ReceiptText3 = value; }
        }

        public string receiptText4
        {
            get { return ReceiptText4; }
            set { ReceiptText4 = value; }
        }

        public string receiptText5
        {
            get { return ReceiptText5; }
            set { ReceiptText5 = value; }
        }

        public string receiptTap
        {
            get { return ReceiptTap; }
            set { ReceiptTap = value; }
        }

        public string receiptSign
        {
            get { return ReceiptSign; }
            set { ReceiptSign = value; }
        }

        public string receiptDate
        {
            get { return ReceiptDate; }
            set { ReceiptDate = value; }   
        }

        public string address
        {
            get { return Address; }
            set { Address = value; }
        }

        public string zipcode
        {
            get { return Zipcode; }
            set { Zipcode = value; }
        }

        /// <summary>
        /// 设置详情表宽度比例，如{1f,2f,2f,1f,1f}
        /// </summary>
        public float[] setTableWidth
        {
            get { return DetailTableWidth; }
            set { DetailTableWidth = value; }
        }

        public int blocks
        {
            get { return Blocks; }
            set { Blocks = value; }
        }

        /// <summary>
        /// 1=auto
        /// 2=manual
        /// 3=hide
        /// </summary>
        public int forwardingType
        {
            get { return ForwardingType; }
            set { ForwardingType = value; }
        }
        #endregion
        public bool CreatePDF()
        {
            var pdf = new Document();//创建pdf
            PdfPTable table;
            PdfPCell cell;
            Paragraph paragraph;
            Chunk chunk;
            BaseFont basefont;
            if (File.Exists(@"C:\Windows\Fonts\simsun.ttc")) {
                basefont = BaseFont.CreateFont(@"C:\Windows\Fonts\simsun.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            }
            else
            {
                basefont = null;
                System.Windows.Forms.MessageBox.Show("Font Error");
                return false;
            }
            Font font_normal, font_bold, font;
            font_normal = new Font(basefont, 14);
            font_bold = new Font(basefont, 16);
            font_bold.SetStyle("bold");
            Image HorizontalLine;

            PdfWriter.GetInstance(pdf, new FileStream(FilePath, FileMode.Create));
            pdf.Open();//开始写入页面

            switch (Blocks)
            {
                case 2: break;//二栏
                case 3://三栏                    
                    PageCount_temp = details.Count-1;
                    for (int paging = 0; paging < PageCount_temp; paging++)
                    {//页面内容

                        //标题
                        font = new Font(basefont, 24);
                        chunk = new Chunk(Title, font);
                        paragraph = new Paragraph(chunk);
                        paragraph.Alignment = Element.ALIGN_CENTER;
                        table = new PdfPTable(1);
                        cell = new PdfPCell(paragraph);
                        cell.BorderWidth = 0;
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.PaddingTop = TitlePaddingTop;
                        cell.PaddingTop = 10f;
                        table.AddCell(cell);
                        pdf.Add(table);

                        //第一个表
                        table = new PdfPTable(6);
                        table.TotalWidth = pdf.PageSize.Width - 120f;
                        table.LockedWidth = true;
                        //接收单位
                        switch (ForwardingType)
                        {
                            case 1://自动匹配
                                Forwarding = "自动匹配：";
                                break;
                            case 2://输入
                                //Forwarding = "";
                                break;
                            case 3://隐藏
                                //Forwarding = "";
                                break;
                            default:
                                System.Windows.Forms.MessageBox.Show("ForwardingType Error");

                                return false;
                        }
                        chunk = new Chunk(Forwarding, font_normal);
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.Colspan = 3;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell.BorderWidthBottom = cell.BorderWidthRight = 0;
                        cell.PaddingLeft = 15f;
                        cell.PaddingTop = 12f;
                        table.AddCell(cell);
                        //档案号
                        switch (FileNumType)
                        {
                            case 1://自动匹配
                                FileNum_Edited = FileType + "档字：" + FileNum + "号";
                                break;
                            case 2://流水号
                                string FileNumTemp = FileNum;
                                FileNum = (Convert.ToInt32(FileNumTemp) + paging).ToString("D" + FileNumLength);
                                FileNum_Edited = FileType + "档字：" + FileNum + "号";
                                break;
                            case 3://隐藏
                                FileNum_Edited = "";
                                break;
                            default:
                                System.Windows.Forms.MessageBox.Show("FileNumType Error");

                                return false;
                        }
                        chunk = new Chunk(FileNum_Edited, font_normal);
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.Colspan = 3;
                        cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cell.BorderWidthLeft = cell.BorderWidthBottom = 0;
                        cell.PaddingRight = 15f;
                        cell.PaddingTop = 12f;
                        table.AddCell(cell);
                        //主要
                        paragraph = new Paragraph();
                        //paragraph.Add("        ");
                        chunk = new Chunk(MainText1, font_normal);
                        paragraph.Add(chunk);
                        Name = Details[paging+1][0];//按照数据匹配，默认Detail[paging+1,0]
                        chunk = new Chunk(Name, font_bold);
                        paragraph.Add(chunk);
                        chunk = new Chunk(MainText2, font_normal);
                        paragraph.Add(chunk);
                        chunk = new Chunk(PeopleCount, font_bold);
                        paragraph.Add(chunk);
                        chunk = new Chunk(MainText3, font_normal);
                        paragraph.Add(chunk);
                        chunk = new Chunk(FileCount1, font_bold);
                        paragraph.Add(chunk);
                        chunk = new Chunk(MainText4, font_normal);
                        paragraph.Add(chunk);
                        chunk = new Chunk(FileCount2, font_bold);
                        paragraph.Add(chunk);
                        chunk = new Chunk(MainText5, font_normal);
                        paragraph.Add(chunk);
                        cell = new PdfPCell(paragraph);
                        cell.Colspan = 6;
                        cell.BorderWidthTop = cell.BorderWidthBottom = 0;
                        cell.PaddingTop = 25f;
                        cell.PaddingLeft = cell.PaddingRight = 15f;
                        cell.PaddingBottom = 10f;
                        table.AddCell(cell);
                        //落款
                        chunk = new Chunk(Sending, font_normal);
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cell.PaddingRight = 30f;
                        cell.Colspan = 6;
                        cell.BorderWidthTop = cell.BorderWidthBottom = 0;
                        table.AddCell(cell);
                        //日期
                        chunk = new Chunk(Date, font_normal);
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cell.PaddingRight = 30f;
                        cell.Colspan = 6;
                        cell.BorderWidthTop = cell.BorderWidthBottom = 0;
                        cell.PaddingBottom = 15f;
                        table.AddCell(cell);
                        //下方详情
                        font = new Font(basefont, 13);
                        PdfPTable tempTable = new PdfPTable(Details[0].Count);
                        for (int count1 = 0; count1 < Details[0].Count; count1++)
                        {
                            //for (int count2 = 0; count2 < Details[0].Count; count2++)
                            //{
                                
                                    font.SetStyle("bold");
                                
                                chunk = new Chunk(Details[0][count1], font);
                                cell = new PdfPCell(new Phrase(chunk));
                                cell.Padding = 2f;
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                cell.BorderWidthRight = cell.BorderWidthBottom = 0;
                                tempTable.AddCell(cell);
                            //}
                        }
                        for (int count1 = 0; count1 < Details[0].Count; count1++)
                        {
                            //for (int count2 = 0; count2 < Details[0].Count; count2++)
                            //{
                            
                                font = new Font(basefont, 12);
                                font.SetStyle("normal");
                            
                            chunk = new Chunk(Details[paging+1][count1], font);
                            cell = new PdfPCell(new Phrase(chunk));
                            cell.Padding = 2f;
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                            cell.BorderWidthRight = cell.BorderWidthBottom = 0;
                            tempTable.AddCell(cell);
                            //}
                        }
                        cell = new PdfPCell(tempTable);
                        cell.Padding = 0f;
                        cell.Colspan = 6;
                        cell.BorderWidthTop = cell.BorderWidthLeft = 0;
                        tempTable.SetWidths(DetailTableWidth);
                        table.AddCell(cell);
                        table.SpacingAfter = table.SpacingBefore = 20f;
                        pdf.Add(table);

                        //分割线
                        HorizontalLine = Image.GetInstance(@"./line.bmp");
                        HorizontalLine.SetAbsolutePosition(0, pdf.PageSize.Height - table.TotalHeight - pdf.TopMargin - 76f);
                        pdf.Add(HorizontalLine);

                        //第二个表
                        pdf.Add(table);

                        //分割线
                        HorizontalLine = Image.GetInstance(@"./solidline.bmp");
                        HorizontalLine.SetAbsolutePosition(0, pdf.PageSize.Height - table.TotalHeight * 2 - pdf.TopMargin - 116f);
                        pdf.Add(HorizontalLine);

                        //第三个表
                        table = new PdfPTable(2);
                        table.TotalWidth = pdf.PageSize.Width - 120f;
                        table.LockedWidth = true;
                        table.SetWidths(new float[] { 1f, 11f });
                        //左
                        font = new Font(basefont, 20);
                        font.SetStyle("bold");
                        chunk = new Chunk(Receipt, font);
                        paragraph = new Paragraph();
                        paragraph.Add(chunk);
                        paragraph.Alignment = Element.ALIGN_CENTER;
                        cell = new PdfPCell(paragraph);
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        table.AddCell(cell);
                        //右
                        tempTable = new PdfPTable(4);
                        //回执接收单位
                        chunk = new Chunk(ReceiptForward, font_normal);
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.Colspan = 2;
                        cell.BorderWidth = 0;
                        cell.PaddingTop = 12f;
                        cell.PaddingLeft = 10f;
                        tempTable.AddCell(cell);
                        //回执档案号
                        chunk = new Chunk(FileNum_Edited, font_normal);
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.Colspan = 2;
                        cell.BorderWidth = 0;
                        cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cell.PaddingTop = 12f;
                        cell.PaddingRight = 10f;
                        tempTable.AddCell(cell);
                        //主要
                        paragraph = new Paragraph();
                        //paragraph.Add("        ");
                        chunk = new Chunk(ReceiptText1, font_normal);
                        paragraph.Add(chunk);
                        if (FileNumType != 3)
                        {
                            chunk = new Chunk(" ", font_normal);
                            paragraph.Add(chunk);
                            chunk = new Chunk(FileNum_Edited, font_bold);
                            paragraph.Add(chunk);
                            chunk = new Chunk(" ", font_normal);
                            paragraph.Add(chunk);
                        }
                        chunk = new Chunk(ReceiptText2, font_normal);
                        paragraph.Add(chunk);
                        chunk = new Chunk(Name, font_bold);
                        paragraph.Add(chunk);
                        chunk = new Chunk(ReceiptText3, font_normal);
                        paragraph.Add(chunk);
                        chunk = new Chunk(FileCount1, font_bold);
                        paragraph.Add(chunk);
                        chunk = new Chunk(MainText4, font_normal);
                        paragraph.Add(chunk);
                        chunk = new Chunk(FileCount2, font_bold);
                        paragraph.Add(chunk);
                        chunk = new Chunk(ReceiptText4, font_normal);
                        paragraph.Add(chunk);
                        cell = new PdfPCell(paragraph);
                        cell.Colspan = 4;
                        cell.BorderWidth = 0;
                        cell.Padding = 15f;
                        tempTable.AddCell(cell);
                        //此复
                        chunk = new Chunk(ReceiptText5, font_normal);
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.Colspan = 3;
                        cell.BorderWidth = 0;
                        cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                        tempTable.AddCell(cell);
                        chunk = new Chunk(" ");
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.BorderWidth = 0;
                        cell.PaddingTop = 15f;
                        tempTable.AddCell(cell);
                        //盖章
                        chunk = new Chunk(ReceiptTap, font_normal);
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.Colspan = 2;
                        cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cell.BorderWidth = 0;
                        cell.PaddingTop = 5f;
                        tempTable.AddCell(cell);
                        //签字
                        chunk = new Chunk(ReceiptSign, font_normal);
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.Colspan = 2;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell.BorderWidth = 0;
                        cell.PaddingTop = 5f;
                        tempTable.AddCell(cell);
                        //年月日
                        chunk = new Chunk(ReceiptDate, font_normal);
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.Colspan = 3;
                        cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                        cell.BorderWidth = 0;
                        cell.PaddingBottom = 15f;
                        tempTable.AddCell(cell);
                        chunk = new Chunk(" ");
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.BorderWidth = 0;
                        cell.PaddingTop = 15f;
                        tempTable.AddCell(cell);
                        //添加子表至宿主表
                        cell = new PdfPCell(tempTable);
                        cell.Padding = 0f;
                        table.AddCell(cell);
                        //尾部地址
                        font = new Font(basefont, 12);
                        chunk = new Chunk(Address, font);
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.Colspan = 2;
                        cell.BorderWidth = 0;
                        table.AddCell(cell);
                        //邮编及时间
                        chunk = new Chunk(Zipcode, font);
                        cell = new PdfPCell(new Phrase(chunk));
                        tempTable = new PdfPTable(2);
                        cell.BorderWidth = 0;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        tempTable.AddCell(cell);
                        chunk = new Chunk(Date, font);
                        cell = new PdfPCell(new Phrase(chunk));
                        cell.BorderWidth = 0;
                        cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                        tempTable.AddCell(cell);
                        cell = new PdfPCell(tempTable);
                        cell.Colspan = 2;
                        cell.BorderWidth = 0;
                        table.AddCell(cell);
                        table.SpacingBefore = 20f;
                        pdf.Add(table);
                        if (paging!=PageCount_temp-1)
                        {
                            pdf.NewPage();
                        }
                    }
                    break;
                default:
                    pdf.Close();
                    System.Windows.Forms.MessageBox.Show("Blocks Error");

                    return false;
                    
            }

            pdf.Close();

            return true;
        }

        /// <summary>
        /// 设置档案号格式:{"{1}","{2}"}
        /// </summary>
        /// <param name="Setting">{1}=1,2,3;{2}=format</param>
        /// <returns></returns>
        public bool SetFileNum(string[] Setting)
        {
            switch (Setting[0])
            {
                case "1":
                    FileNumType = 1;
                    break;
                case "2":
                    FileNumType = 2;
                    FileNumLength = Convert.ToInt32(Setting[1]);
                    break;
                case "3":
                    FileNumType = 3;
                    break;
                default:
                    break;
            }
            return true;
        }

       
    }
}
