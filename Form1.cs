using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CommentToAttendence
    {
    public partial class Form1 : Form
        {

        string attachment_filename;
        string sheetname;
        string date = DateTime.Today.ToShortDateString();
        string email_body;
        string initialDirectory;
        string currentDirectory;
        string text;
        bool isPrinted = false;
        bool isPrintedwhole = false;

        //HashSet<string> id = new HashSet<string>();
        SortedSet<string> id = new SortedSet<string>();
        //SortedSet<string> name = new SortedSet<string>();
        DataTable IDTbl = new DataTable();
        DataTable StudentTbl = new DataTable();
        

        ExcelCreator excelCreator = new ExcelCreator();
        ExcelCreatorRaw excelCreator1 = new ExcelCreatorRaw();
        MailCreator mail = new MailCreator();
        DB dbAccess = new DB();

        public Form1()
            {
            InitializeComponent();
            }

        private void Form1_Load(object sender, EventArgs e)
            {
            saveFileDialog1.Title = "Save Attendance Excel File";
            saveFileDialog1.Filter = "Excel Documents|*.xlsx";

            initialDirectory =  Directory.GetCurrentDirectory();
            attachment_filename = "Attandence_" + date + ".xlsx";
            email_body = sheetname + "   Date: " + date;

            IDTbl.Columns.Add("SL", typeof(int));
            //IDTbl.Columns.Add("Name");
            IDTbl.Columns.Add("ID");

            label5.Text = date;
            //this.ActiveControl = tabPage1;
            //tabControl1.SelectTab(0);

            showstudentTbl();
            dataGridView2.DataSource = StudentTbl;
            dataGridView2.Font = new Font("HP Simplified", 9);
            // dataGridView2.DefaultCellStyle.Font = new Font("HP Simplified", 9);

            button5.Enabled = isPrinted;
            button6.Enabled = isPrintedwhole;
            }

        ///Convert Btn
        private void button1_Click(object sender, EventArgs e)
            {
            refreshSheet();

            //commentToName();
            commentToID();
            insertIDTbl();
            dataGridView1.DataSource = IDTbl;
            dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
            label6.Text = IDTbl.Rows.Count.ToString();
            }
        
        ///Print Btn
        private void button2_Click(object sender, EventArgs e)
            {

            sheetname = textBox3.Text;
            string filename = sheetname+"_Attendance_" + date;

            saveFileDialog1.FileName = filename;
            saveFileDialog1.InitialDirectory = initialDirectory;

            try
                {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                    string backcolor_cell = "#F0FFFF";
                    var textcolor = System.Drawing.Color.DarkBlue;

                    currentDirectory = saveFileDialog1.FileName;
                    MessageBox.Show(currentDirectory);

                    excelCreator.DataTableToExcel(IDTbl, sheetname, sheetname + " Attendance", date);
                    excelCreator.FormattingExcelCells(ExcelCreator.excelCellrange, backcolor_cell, textcolor, true);
                    //savefile();

                    excelCreator.excelSaveAsFile(currentDirectory);
                    excelCreator.ProcessTermination();
                    excelCreator.ReleaseAllComObjects();

                    isPrinted = true;
                    button5.Enabled = isPrinted;
                    }

                }
            catch
                {
                MessageBox.Show("Course Code/Lecture No. is not given");
                }

            /*
            try 
                {
                string backcolor_cell = "#F0FFFF";
                var textcolor = System.Drawing.Color.DarkBlue;

                sheetname = textBox3.Text;
                string filepath = filedirectory;
                string filename = "Attendance_" + date + ".xlsx";
                excelCreator.DataTableToExcel(IDTbl, sheetname, sheetname + " Attendance", date);
                excelCreator.FormattingExcelCells(ExcelCreator.excelCellrange, backcolor_cell, textcolor, true);
                excelCreator.excelSaveAs(filepath, filename);     
                excelCreator.ProcessTermination();
                excelCreator.ReleaseAllComObjects();
                }
            catch
                {
                MessageBox.Show("Course Code/Lecture No. is not given");
                }

            */

            }

        ///Clear Btn
        private void button3_Click(object sender, EventArgs e)
            {
            textBox1.Text = "";
            }

        ///Calculate Btn
        private void button4_Click(object sender, EventArgs e)
            {
            bool isupdated = false;
            int rowID = IDTbl.Rows.Count;
            int rowST = StudentTbl.Rows.Count;
            for (int i=0;i< rowID; i++)
                {
                for (int j = 0; j < rowST; j++)
                    {
                        if(IDTbl.Rows[i].Field<string>("ID") == StudentTbl.Rows[j].Field<string>("Student_ID"))
                        {
                        string id = IDTbl.Rows[i].Field<string>("ID");
                        string query = " UPDATE Attandence_Count SET Student_Attan_Count = Student_Attan_Count+1  WHERE Student_ID =@id ";
                        SqlCommand updateCommand = new SqlCommand(query);
                        updateCommand.Parameters.AddWithValue("@id", id);
                        int row = dbAccess.executeQuery(updateCommand);
                        dbAccess.closeConn();
                        if(row>0) isupdated=true; 
                        }
                    }
                }

            if (isupdated) MessageBox.Show("Total " + rowID + " student's attendance calculated!");
            else MessageBox.Show("No student found");
            refreshSheet();
            showstudentTbl();
            }

        ///Mail sheet btn
        private void button5_Click(object sender, EventArgs e)
            {
            try
               {
                sheetname = textBox3.Text;
                string recipient_email_to_other = textBox2.Text.ToLower().Trim();
                //string filepath = initialDirectory;
                //string filename = "Attendance_" + date + ".xlsx";
                string subject = mail.emailSubject + sheetname;
                string body = mail.email_body_header + email_body + mail.email_body_footer + mail.soft_detail;

                //string backcolor_cell = "#F0FFFF";
                var textcolor = System.Drawing.Color.DarkBlue;

                MemoryStream attachment = new MemoryStream();
                //excelCreator.DataTableToExcel(IDTbl, sheetname, sheetname + " Attendance", date);
                //excelCreator.FormattingExcelCells(ExcelCreator.excelCellrange, backcolor_cell, textcolor, true);
                //excelCreator.excelSaveAs(filepath, filename);

                //attachment = excelCreator.ReadingfromFileConvertingToMemoryStream(filepath, filename);
                attachment = excelCreator.ReadingfromFileConvertingToMemoryStream(currentDirectory);
                mail.SendEmail(mail.defalultSender, recipient_email_to_other, subject, body, attachment, attachment_filename);
                attachment.Close();
                attachment.Dispose();
                excelCreator.ProcessTermination();
                excelCreator.ReleaseAllComObjects();
                MessageBox.Show("Attendance Sheet Emailed Successfully to\n" + recipient_email_to_other);
            }
        catch (Exception ex)
            {
            MessageBox.Show(ex.Message);
            }

            }

        ///Attandence Count print Btn
        private void button7_Click(object sender, EventArgs e)
            {

            sheetname = textBox3.Text;
            string filename = sheetname +"Lecture No. "+ textBox4.Text+ "_Attendance_" + date;
            saveFileDialog1.FileName = filename;
            saveFileDialog1.InitialDirectory = initialDirectory;

                try
                {
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                        string backcolor_cell = "#F0FFFF";
                        var textcolor = System.Drawing.Color.DarkBlue;

                        currentDirectory = saveFileDialog1.FileName;

                        excelCreator.DataTableToExcel(StudentTbl, sheetname, sheetname + " Attendance", date);
                        excelCreator.FormattingExcelCells(ExcelCreator.excelCellrange, backcolor_cell, textcolor, true);
                        //excelCreator.excelSaveAs(filepath, filename);
                        excelCreator.excelSaveAsFile(currentDirectory);

                        excelCreator.ProcessTermination();
                        excelCreator.ReleaseAllComObjects();

                        isPrintedwhole = true;
                        button6.Enabled = isPrintedwhole;
                    }
                }
            catch 
                {
                MessageBox.Show("Course Code/Lecture No. is not given");
                }
            }

        ///Attandence Count mail Btn
        private void button6_Click(object sender, EventArgs e)
            {
            try
                {
                sheetname = textBox3.Text;
                string recipient_email_to_other = textBox5.Text.ToLower().Trim();
                //string filepath = initialDirectory;
                //string filename = "Attendance_" + date + ".xlsx";
                string subject = mail.emailSubject + sheetname;
                string body = mail.email_body_header + email_body + mail.email_body_footer + mail.soft_detail;

                //string backcolor_cell = "#F0FFFF";
                var textcolor = System.Drawing.Color.DarkBlue;

                MemoryStream attachment = new MemoryStream();
                //excelCreator.DataTableToExcel(IDTbl, sheetname, sheetname + " Attendance", date);
                //excelCreator.FormattingExcelCells(ExcelCreator.excelCellrange, backcolor_cell, textcolor, true);
                //excelCreator.excelSaveAs(filepath, filename);

                //attachment = excelCreator.ReadingfromFileConvertingToMemoryStream(filepath, filename);
                attachment = excelCreator.ReadingfromFileConvertingToMemoryStream(currentDirectory);
                mail.SendEmail(mail.defalultSender, recipient_email_to_other, subject, body, attachment, attachment_filename);
                attachment.Close();
                attachment.Dispose();
                excelCreator.ProcessTermination();
                excelCreator.ReleaseAllComObjects();
                MessageBox.Show("Attendance Sheet Emailed Successfully to\n" + recipient_email_to_other);
                }
            catch (Exception ex)
                {
                MessageBox.Show(ex.Message);
                }
            }

        void commentToID()
            {
            text = textBox1.Text.Trim();
            string w = "";
            for (int i = 0; i < text.Length; i++)
                {
                char c = text[i];

                if (c >= 48 && c <= 57)
                    {
                    w += c;
                    }
                else if (w != "")
                    {
                    if(w.Length==7) id.Add(w);
                    w = "";
                    }

                }

            if (w != "")
                {
                if (w.Length == 7) id.Add(w);
                    w = "";
                }

            }
        /*
        void commentToName()
            {
            text = textBox1.Text.Trim();
            string w = "1234567";
            string n = "";
            for (int i = 0; i < text.Length; i++)
                {
                
                char c = text[i];

                if (c >= 48 && c <= 57)
                    {
                    w += c;
                    }
                else if (w != "")
                    {
                    w = "";
                    }

                if (c!='_' && w.Length==7)
                    {
                    n += c;
                    }
                else if (n != "")
                    {
                    name.Add(n);
                    n = "";
                    }

                }

            if (w != "")
                {
                w = "";
                }

            if (n != "")
                {
                 name.Add(n);
                n = "";
                }

            }
        */

        void insertIDTbl()
            {
            for(int i = 0; i < id.Count; i++)
                {
                    IDTbl.Rows.Add(i+1, id.ElementAt(i));
                }
           
            }

        void refreshSheet()
            {

            //IDTbl.Columns.Clear();
            IDTbl.Rows.Clear();
            //IDTbl.Clear();       
            }

        void showstudentTbl()
            {
            StudentTbl.Rows.Clear();
            string query = " SELECT * FROM Attandence_Count ";
            dbAccess.readDatathroughAdapter(query, StudentTbl);
            }

        private void button1_MouseLeave(object sender, EventArgs e)
            {
            button1.ForeColor = Color.Crimson;
            button1.BackColor = Color.LemonChiffon;
            }

        private void button1_MouseHover(object sender, EventArgs e)
            {
            button1.BackColor = Color.Crimson;
            button1.ForeColor = Color.LemonChiffon;
            }

        private void button2_MouseLeave(object sender, EventArgs e)
            {
            button2.ForeColor = Color.Green;
            button2.BackColor = Color.Azure;
            }

        private void button2_MouseHover(object sender, EventArgs e)
            {
            button2.BackColor = Color.Green;
            button2.ForeColor = Color.Azure;
            }

        /*
        void savefile()
            {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Save Excel Files";
            saveFileDialog1.Filter = "Exel files (*.xlsx)|*.xlsx";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                filedirectory = saveFileDialog1.FileName;
                excelCreator1.excelSaveAsFile(filedirectory);
                excelCreator1.ProcessTermination();
                excelCreator1.ReleaseAllComObjects();
                }
            }
        */


        }
    }
