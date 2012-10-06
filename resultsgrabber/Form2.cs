using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
					

namespace resultsgrabber
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int start = Convert.ToInt32(textBox1.Text);
            int end = Convert.ToInt32(textBox2.Text);
            int i = start;
            int j = end;
            double critical = Convert.ToDouble(textBox3.Text);
            String url = "http://www.annauniv.edu/cgi-bin/tancet2011/tancet2011res.pl";
            HttpWebResponse response;
            String responseString;
            Stream streamResponse;
            StreamReader streamReader;
            byte[] byteArray;
            int BUFF = 1024;
            char[] readBuffer = new char[BUFF];
            //StringBuilder postString = new StringBuilder("regno=xxxxxx&B1=Get+Marks");
            StringBuilder postString = new StringBuilder("tan2011rno=xxxxxxxx");
            Stream postStream;
            var myExcelApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            dynamic myExcelWorkbook = myExcelApp.Workbooks.Add();
            Excel.Worksheet myExcelWorksheet = myExcelWorkbook.ActiveSheet;
            myExcelApp.Visible = true;
            myExcelWorksheet.get_Range("A1", misValue).Formula = "Roll No";
            myExcelWorksheet.get_Range("B1", misValue).Formula = "Name";
            myExcelWorksheet.get_Range("C1", misValue).Formula = "marks";
            int count = 2;
            for (; i <= j; i++)
            {
                try
                {
                    HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(url);
                    myRequest.Method = "POST";
                    myRequest.ContentType = "application/x-www-form-urlencoded";
                    //myRequest.ContentLength = 25;
                    //postString.Remove(6, 6);
                    //postString.Insert(6, i);
                    myRequest.ContentLength = 19;
                    postString.Remove(11,8);
                    postString.Insert(11, i);
                    byteArray = Encoding.UTF8.GetBytes(postString.ToString());

                    postStream = myRequest.GetRequestStream();
                    postStream.Write(byteArray, 0, byteArray.Length);
                    response = (HttpWebResponse)myRequest.GetResponse();
                    streamResponse = response.GetResponseStream();
                    streamReader = new StreamReader(streamResponse);
                    responseString = streamReader.ReadToEnd();
                    streamReader.Close();
                    response.Close();
                    streamResponse.Close();




                    try
                    {

                        int index = 0;
                        myExcelWorksheet.get_Range("A" + count.ToString(), misValue).Formula = i.ToString();
                        if(responseString.IndexOf("Sorry")>-1)
                        {
                            continue;
                        }
                        String marks = getMark(responseString, ref index);

                        if (marks.IndexOf("ABS") < 0 && marks.IndexOf("--") < 0 && Convert.ToDouble(marks) >= critical)
                        {
                            myExcelWorksheet.get_Range("B" + count.ToString(), misValue).Formula = getName(responseString, ref index);
                            myExcelWorksheet.get_Range("C" + count.ToString(), misValue).Formula = marks;
                            count++;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }


            }
            MessageBox.Show("finished");
            //myExcelWorkbook.saveAs("12Results"+textBox1.Text+"-"+textBox2.Text+".xsl", AccessMode: Excel.XlSaveAsAccessMode.xlShared);
            myExcelWorkbook.saveAs("mca_" + textBox1.Text + "-" + textBox2.Text + " _.xsl", AccessMode: Excel.XlSaveAsAccessMode.xlShared);
        }

        private String getMark(string html, ref int start)
        {
            int end;
          
            StringBuilder name = new StringBuilder();
            start = html.IndexOf("M.C.A", 200);
            start = html.IndexOf("&nbsp;", start)+12;
            end = html.IndexOf("<",start);
                       

            name.Append(html, start, end-start);
            start = end;
            
            return name.ToString();
        }
        private String getName(string html, ref int start)
        {

            int end;
            StringBuilder name = new StringBuilder();
            start = html.IndexOf("Name", 200);
            start = html.IndexOf("\">", start);
            start = html.IndexOf("24", start)+4;
            end = html.IndexOf("</", start);

            name.Append(html, start, end - start);
            start = end;
            return name.ToString();
        }

    }

}
