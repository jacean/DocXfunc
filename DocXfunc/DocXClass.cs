using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Drawing;
using Novacode;
using Image = Novacode.Image;

namespace DocXfunc
{
    class DocXClass
    {
        private bool isExist(string file)
        {
            return (!File.Exists(file));
           
        }
        public static void createDoc()
        {
            createDoc(Application.StartupPath+@"\DocX\","newDocx.docx");
        }

        public static void createDoc(string Path)
        {
            
            createDoc(@Path, "newDocx.docx");
        }
        public static void createDoc(string Path,String Name)
        {
            if (!Directory.Exists(Path))
            {
                Directory.CreateDirectory(Path);
            }
            if (Name.Split('.')[1] != "docx")
            {
                Name = Name.Split('.')[0] + ".docx";
            }
            string docx = @Path + "\\" + Name;
            using (DocX document = DocX.Create(docx))
            {
                //// Insert a Paragraph into this document.
                //Paragraph p = document.InsertParagraph();

                //// Append some text and add formatting.
                //p.Append("Hello World!^011Hello World!")
                //.Font(new FontFamily("Times New Roman"))
                //.FontSize(32)
                //.Color(Color.Blue)
                //.Bold();
        

                document.Save();
                //Console.WriteLine("\tCreated: docs\\Hello World.docx\n");
            }
        }
        public static void setHeaderFooter(string docx,string headerstr,string footerstr)
        {
            if (!File.Exists(docx))
            {
                MessageBox.Show(docx+"文件不存在");
                return;
            }
            DocX document = DocX.Load(docx);
            
            document.AddHeaders();
            document.AddFooters();
            // Force the first page to have a different Header and Footer.
            document.DifferentFirstPage = false;
            // Force odd & even pages to have different Headers and Footers.
            document.DifferentOddAndEvenPages = false;
            #region 设置第一页、奇偶页眉页脚
            // Get the first, odd and even Headers for this document.
            //Header header_first = document.Headers.first;
            //Header header_odd = document.Headers.odd;
            //Header header_even = document.Headers.even;

            //// Get the first, odd and even Footer for this document.
            //Footer footer_first = document.Footers.first;
            //Footer footer_odd = document.Footers.odd;
            //Footer footer_even = document.Footers.even;

            // Insert a Paragraph into the first Header.
            //Paragraph p0 = header_first.InsertParagraph();
            //p0.Append("Hello First Header.").Bold();



            // Insert a Paragraph into the odd Header.
            //Paragraph p1 = header_odd.InsertParagraph();
            //p1.Append("Hello Odd Header.").Bold();


            //// Insert a Paragraph into the even Header.
            //Paragraph p2 = header_even.InsertParagraph();
            //p2.Append("Hello Even Header.").Bold();

            //// Insert a Paragraph into the first Footer.
            //Paragraph p3 = footer_first.InsertParagraph();
            //p3.Append("Hello First Footer.").Bold();

            //// Insert a Paragraph into the odd Footer.
            //Paragraph p4 = footer_odd.InsertParagraph();
            //p4.Append("Hello Odd Footer.").Bold();

            //// Insert a Paragraph into the even Header.
            //Paragraph p5 = footer_even.InsertParagraph();
            //p5.Append("Hello Even Footer.").Bold();
            #endregion

            #region 插入新页、节
            // Insert a Paragraph into the document.
            //Paragraph p6 = document.InsertParagraph();
            //p6.AppendLine("Hello First page.");

            //// Create a second page to show that the first page has its own header and footer.
            //p6.InsertPageBreakAfterSelf();

            //// Insert a Paragraph after the page break.
            //Paragraph p7 = document.InsertParagraph();
            //p7.AppendLine("Hello Second page.");

            //// Create a third page to show that even and odd pages have different headers and footers.
            //p7.InsertPageBreakAfterSelf();

            //// Insert a Paragraph after the page break.
            //Paragraph p8 = document.InsertParagraph();
            //p8.AppendLine("Hello Third page.");

            ////Insert a next page break, which is a section break combined with a page break
            //document.InsertSectionPageBreak();

            ////Insert a paragraph after the "Next" page break
            //Paragraph p9 = document.InsertParagraph();
            //p9.Append("Next page section break.");

            ////Insert a continuous section break
            //document.InsertSection();

            //Create a paragraph in the new section
            //var p10 = document.InsertParagraph();
            //p10.Append("Continuous section paragraph.");
            #endregion
            Header header = document.Headers.odd;
            //header.Tables.First().SetBorder(TableBorderType.Bottom, new Border(Novacode.BorderStyle.Tcbs_single, BorderSize.one, 1, Color.Black));
            Paragraph p_header = header.Paragraphs.First();
            
            p_header.Append(headerstr);//在此处设置格式
            p_header.Alignment = Alignment.center;
            //试图给页眉加下边框，然而失败了
            //header.Paragraphs[0].Remove(false);
            //Table t = header.InsertTable(1,1);
            
            //t.Rows[0].Cells[0].Paragraphs.First().InsertText(headerstr);
            //t.Rows[0].Cells[0].Paragraphs.First().Alignment = Alignment.center;
            //t.Rows[0].Cells[0].SetBorder(TableCellBorderType.Bottom, new Border(Novacode.BorderStyle.Tcbs_single, BorderSize.one, 1, Color.Black));

          //document.Headers.even = header;
          //document.Headers.odd = header;
          Footer footer = document.Footers.odd;
            Paragraph p_footer = footer.Paragraphs.First();
            p_footer.Append(footerstr);
            p_footer.Alignment = Alignment.center;
            //document.Footers.even = footer;
            //document.Footers.odd = footer;
           
            document.Save();

        }

        public static void addParagraph(string docx,string content)
        {
            if (!File.Exists(docx))
            {
                MessageBox.Show(docx + "文件不存在");
                return;
            }
            DocX document = DocX.Load(docx);
            FontDialog fd = new FontDialog();
            if(fd.ShowDialog()==DialogResult.OK)
            {
                Font font = fd.Font;
               
                Paragraph newP = document.InsertParagraph();
                newP.Append(content)
                    .Font(font.FontFamily)
                    .FontSize(font.Size);
            }
            

            document.Save();
                
        }
        public static void addText(string docx, string content)
        {
            if (!File.Exists(docx))
            {
                MessageBox.Show(docx + "文件不存在");
                return;
            }
            DocX document = DocX.Load(docx);
            Paragraph endP = document.Paragraphs[document.Paragraphs.Count - 1];
            FontDialog fd = new FontDialog();
            if (fd.ShowDialog() == DialogResult.OK)
            {
                Font font = fd.Font;
                endP.Append(content)
                    .Font(font.FontFamily)
                    .FontSize(font.Size);
            }
            endP.Append(content);
            document.Save();
        }

        public static void addTable(string docx,int row,int col,string[] colHeader)
        {
            if (!File.Exists(docx))
            {
                MessageBox.Show(docx + "文件不存在");
                return;
            }
            if(col>colHeader.Length)
            {
                MessageBox.Show("列标题少于给定的列数目");
                return;
            }
            using (DocX document = DocX.Load(docx))
            {
                // Create a new Table with 2 coloumns and 3 rows.
                Table newTable = document.InsertTable(row,col);

                // Set the design of this Table.
                //newTable.Design = TableDesign.Custom;//传统样式，但是没有表格线
                //newTable.Design = TableDesign.TableNormal;//传统样式，但是没有表格线
                newTable.Design = TableDesign.TableGrid;//传统样式，有了表格线

                // Set the coloumn names.
                for (int i = 0; i < col; i++)
                {
                    newTable.Rows[0].Cells[i].Paragraphs.First().InsertText(colHeader[i], false);
                }
                
                               
                document.Save();
            }// Release this document from memory.
        }
        public static DocX getDocx(string docx)
        {
            
            DocX document = DocX.Load(docx);
            
                return document;
           
        }
        public static void saveTable(ref DocX document)
        {
            document.Save();
        }
        public static Table[] getTables( DocX document)
        {
            
                List<Table> ts = new List<Table>();
                ts = null;
                ts = document.Tables;
                return ts.ToArray<Table>();
              
            
        }
        public static void setCellvalue(  Table t,int rowindex,int colindex,string value)
        {
            t.Rows[rowindex].Cells[colindex].Paragraphs.First().InsertText(value, false);
        }
        public static void setTablestyle(Table t,string style,string value)
        {

        }
      
        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="t"></param>
        /// <param name="isRows">true则合并同行不同列，false则合并同列不同行</param>
        /// <param name="index">合并行或列的索引</param>
        /// <param name="startIndex">开始的索引</param>
        /// <param name="endIndex">结束的索引</param>
        public static void mergeCells(Table t,bool isRows,int index,int startIndex,int endIndex)
        {
            if (isRows)
            {
                t.Rows[index].MergeCells(startIndex, endIndex);
            }
            else
            {
                t.MergeCellsInColumn(index, startIndex, endIndex);
            }
        }

        public static void addNewpage(string docx)
        {
            if (!File.Exists(docx))
            {
                MessageBox.Show(docx + "文件不存在");
                return;
            }

            using (DocX document = DocX.Load(docx))
            {
                document.InsertParagraph().InsertPageBreakAfterSelf();
                document.Save();
            }
        }
    }
}
