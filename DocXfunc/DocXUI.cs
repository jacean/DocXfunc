using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Novacode;
using Image = Novacode.Image;

namespace DocXfunc
{
    public partial class DocXUI : Form
    {
        public DocXUI()
        {
            InitializeComponent();
        }
        string file = "";

        private void DocXUI_Load(object sender, EventArgs e)
        {
            file = "F:\\example.docx";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DocXClass.createDoc("F:\\","example.doc");
            DocXClass.setHeaderFooter("F:\\example.docx","header","footer");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DocXClass.addParagraph("F:\\example.docx", textBox1.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DocXClass.addText("F:\\example.docx",textBox1.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DocXClass.addTable(file,3,3,new string[] { "1","2","3"});
            DocX document = DocXClass.getDocx(file);
            Table t= DocXClass.getTables( document)[0];
            //document.Save();
            DocXClass.setCellvalue( t, 1, 0, "a");
            DocXClass.setCellvalue(  t, 1, 1, "b");
            DocXClass.setCellvalue( t, 1, 2, "c");
            DocXClass.mergeCells( t, true, 2, 0, 2);
            DocXClass.setCellvalue( t,2,0,"merge");
            //document.Save();
            DocXClass.saveTable(ref document);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DocXClass.addNewpage(file);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DocXClass.addPicture(file);
        }
    }
}
