using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;

namespace CommentToAttendence
    {
    public class ExcelCreator
        {      
        public static Microsoft.Office.Interop.Excel.Application excel;
        public static Microsoft.Office.Interop.Excel.Workbook excelworkBook;
        public static Microsoft.Office.Interop.Excel.Worksheet excelSheet;
        public static Microsoft.Office.Interop.Excel.Range excelCellrange;

        public void DataTableToExcel(DataTable dataTable, string sheet_name, string sheet_main_header, string sheet_sub_header)
            {
            ///Creation of Excel objects
            excel = new Microsoft.Office.Interop.Excel.Application();
            // for making Excel visible  
            excel.Visible = false;
            excel.DisplayAlerts = false;
            // Creation a new Workbook  
            excelworkBook = excel.Workbooks.Add(Type.Missing);
            // Workk sheet  
            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
            excelSheet.Name = sheet_name;
             
            ///Writing to Excel file
            excelSheet.Cells[1, 1] = sheet_main_header;
            excelSheet.Cells[2, 1] = sheet_sub_header;
            excelSheet.Cells[1, 1].Font.Bold = true;
            excelSheet.Cells[2, 1].Font.Bold = true;
            excelSheet.Cells[1, 1].Font.Size = 13;
            excelSheet.Cells[2, 1].Font.Size = 13;
            
            //Columns
            for (var i = 0; i < dataTable.Columns.Count; i++)
                {
                excelSheet.Cells[4, i + 1] = dataTable.Columns[i].ColumnName;
                }
            //Rows
            for (var i = 0; i < dataTable.Rows.Count; i++)
                {
                // to do: format datetime values before printing
                for (var j = 0; j < dataTable.Columns.Count; j++)
                    {
                    excelSheet.Cells[i + 5, j + 1] = dataTable.Rows[i][j];
                    }
                }


            ///Working with range and formatting Excel cells
            // now we resize the columns  
            excelCellrange = excelSheet.Range[excelSheet.Cells[4, 1], excelSheet.Cells[dataTable.Rows.Count + 4, dataTable.Columns.Count]];
            excelCellrange.EntireColumn.AutoFit();
            Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            //excelworkBook.Save();
            }

        public void excelSaveAs(string filepath, string filename)
            {
            try
                {
                //excelworkBook.Save();
                //excelSheet.SaveAs(filepath + filename);
                excelworkBook.SaveCopyAs(filepath + filename); // SaveCopyAs(filepath+filename+extension);
                //excelworkBook.Close();
                //excel.Visible = true;
                //excel.Quit();

                }
            catch(Exception ex)
                {
                MessageBox.Show(ex.Message);
                }
            }

        public void excelSaveAsFile(string filepathfilename)
            {
            try
                {
                //excelworkBook.Save();
                //excelworkBook.SaveAs(filepathfilename);
                excelworkBook.SaveCopyAs(filepathfilename); // SaveCopyAs(filepath+filename+extension);
                //excelworkBook.Close();
                //excel.Visible = true;
                //excel.Quit();
                }
            catch (Exception ex)
                {
                MessageBox.Show(ex.Message);
                }
            }

        public byte[] excelToBinary()
            {
            try
                {
                MemoryStream memoryStream = new MemoryStream();
                //excelworkBook.SaveCopyAs(memoryStream);
                excelworkBook.SaveAs(memoryStream);
                //excelworkBook.Close();
                //excel.Quit();
                byte[] bytes = memoryStream.ToArray();
                memoryStream.Close();
                memoryStream.Dispose();
                //excel.Visible = true;
                return bytes;
                }
            finally
                {
                //ReleaseAllComObjects();
                }

            }

        public MemoryStream excelToMemoryStream()
            {
            try
                {
                MemoryStream memoryStream = new MemoryStream();
                excelworkBook.SaveCopyAs(memoryStream);
                //excelworkBook.SaveAs(memoryStream);
                //excelworkBook.Close();
                //excel.Quit();
                //memoryStream.Close();
                memoryStream.Dispose();
                //excel.Visible = true;
                return memoryStream;
                }
            finally
                {
                //ReleaseAllComObjects();
                }

            }

        public MemoryStream excelSaveAsUsingMemoryStream()
            {

            MemoryStream MyMemoryStream = new MemoryStream();
            excelworkBook.SaveCopyAs(MyMemoryStream);
          
            //write to file
            FileStream file = new FileStream("C:\\Users\\HP 840 G1\\Desktop\\Memostrxl.xlsx", FileMode.Create, FileAccess.Write,FileShare.Write);
            MyMemoryStream.WriteTo(file);
            file.Close();
            MyMemoryStream.Close();
            MyMemoryStream.Dispose();
            //excel.Visible = true;
            return MyMemoryStream;
            }

        /*
        public MemoryStream ReadingfromFileConvertingToMemoryStream(string filepath, string filename)
            {
            string fullpathwithexten = filepath + filename;
            string withafullpathwithexten = @fullpathwithexten;
            using (FileStream fileStream = File.Open(withafullpathwithexten,FileMode.Open,FileAccess.ReadWrite,FileShare.ReadWrite))
                {
                MemoryStream memStream = new MemoryStream();
                memStream.SetLength(fileStream.Length);
                fileStream.Read(memStream.GetBuffer(), 0, (int)fileStream.Length);
                return memStream;
                }
            }
        */

        public MemoryStream ReadingfromFileConvertingToMemoryStream(string fullpathwithexten)
            {
            //string withafullpathwithexten = @fullpathwithexten;
            using (FileStream fileStream = File.Open(fullpathwithexten, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                MemoryStream memStream = new MemoryStream();
                memStream.SetLength(fileStream.Length);
                fileStream.Read(memStream.GetBuffer(), 0, (int)fileStream.Length);
                return memStream;
                }
            }

        ///Coloring cells
        public void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbolt)
            {
            range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            if (IsFontbolt == true)
                {
                range.Font.Bold = IsFontbolt;
                }
            //excelworkBook.Close();
            //excel.Visible = true;
            //excel.Quit();
            }

        ///Excel Process Terminating
        //This is must to Terminate the processes otherwise Excel workbook/worksheet operations can't be done 
        public void ProcessTermination()
           {
             foreach (Process proc in System.Diagnostics.Process.GetProcessesByName("Excel"))
               {
                proc.Kill();
               }
            }

        ///Release The COM Objects
        // This is must for the COM objects because it handles unexpected Excel File behavior
        public void ReleaseAllComObjects()
            {
            Marshal.ReleaseComObject(excelCellrange);
            Marshal.ReleaseComObject(excelSheet);
            Marshal.ReleaseComObject(excelworkBook);
            Marshal.ReleaseComObject(excel);
            }

        /*
            private byte[] ConvertDataSetToByteArray(DataTable dataTable)
            {
                byte[] binaryDataResult = null;
                using (MemoryStream memStream = new MemoryStream())
                {
                BinaryFormatter brFormatter = new BinaryFormatter();
                dataSet.RemotingFormat = SerializationFormat.Binary;
                brFormatter.Serialize(memStream, dataTable);
                binaryDataResult = memStream.ToArray();
                }
                return binaryDataResult;
             }
        */


        }
    }
