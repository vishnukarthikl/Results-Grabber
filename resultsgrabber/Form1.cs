﻿
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
namespace resultsgrabber
{


    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private String getName(String html)
        {
            int start;
            int end;
            StringBuilder name = new StringBuilder();
            start=html.IndexOf("<b>", 200)+3;
            end=html.IndexOf("&nbsp",start);
           
            name.Append(html,start,end-start);
            return name.ToString();
        }
        private String getLanguage(string html)
        {
            int start;
            int end;
            StringBuilder name = new StringBuilder();
            start = html.IndexOf("LANGUAGE",300);
            start = html.IndexOf("<b>", start);
            end = html.IndexOf("</b>", start);
            name.Append(html, start + 15,end-(start+15));
            return name.ToString();

 
        }
        private String getEnglish(string html)
        {
            int start;
            int end;
            StringBuilder name = new StringBuilder(); 
            start = html.IndexOf("ENGLISH", 300);
            start = html.IndexOf("<b>", start);
            end = html.IndexOf("</b>", start);
            name.Append(html, start + 15, end - (start + 15));
            return name.ToString();


        }

        private void button1_Click(object sender, EventArgs e)
        {
          
            int start = Convert.ToInt32(textBox1.Text);
            int end = Convert.ToInt32(textBox2.Text);
            int i = start <= end ? start : end;
            int j = start <= end ? end : start;
            int count=2;
            //String url = "http://tnresults.nic.in/hsc/result.asp";
            String url = "http://tnresults.nic.in/dgebase/final.asp";
            HttpWebResponse response;
            String responseString;
            Stream streamResponse;
            StreamReader streamReader;
            byte[] byteArray;
            int BUFF = 1024;
            char[] readBuffer = new char[BUFF];
            //StringBuilder postString = new StringBuilder("regno=xxxxxx&B1=Get+Marks");
            StringBuilder postString = new StringBuilder("etype=S&regno=xxxxxxx&B1=Get+Marks");
            Stream postStream;
            var myExcelApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            dynamic myExcelWorkbook = myExcelApp.Workbooks.Add();
            Excel.Worksheet myExcelWorksheet = myExcelWorkbook.ActiveSheet;
            myExcelApp.Visible = true;
            myExcelWorksheet.get_Range("A1", misValue).Formula = "Roll No";
            myExcelWorksheet.get_Range("B1", misValue).Formula = "Name";
            myExcelWorksheet.get_Range("C1", misValue).Formula = "Language";
            myExcelWorksheet.get_Range("D1", misValue).Formula = "English";
            myExcelWorksheet.get_Range("E1", misValue).Formula = "Maths";
            myExcelWorksheet.get_Range("F1", misValue).Formula = "Science";
            myExcelWorksheet.get_Range("G1", misValue).Formula = "Social Science";
            myExcelWorksheet.get_Range("H1", misValue).Formula = "Total";
            myExcelWorksheet.get_Range("I1", misValue).Formula = "Result";
            //myExcelWorksheet.get_Range("A1", misValue).Formula = "Name";
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
                myRequest.ContentLength = 34;
                postString.Remove(14,7);
                postString.Insert(14, i);
                byteArray= Encoding.UTF8.GetBytes(postString.ToString());
                
                postStream = myRequest.GetRequestStream();
                postStream.Write(byteArray, 0, byteArray.Length);
                response= (HttpWebResponse)myRequest.GetResponse();
                streamResponse= response.GetResponseStream();
                streamReader = new StreamReader(streamResponse);
                responseString = streamReader.ReadToEnd();
                streamReader.Close();
                response.Close();
                streamResponse.Close();

                            
               

                try
                {

                    int index=0;
                    myExcelWorksheet.get_Range("A" + count.ToString(), misValue).Formula = i.ToString();
                    myExcelWorksheet.get_Range("B" + count.ToString(), misValue).Formula = getName10(responseString, ref index);
                    myExcelWorksheet.get_Range("C" + count.ToString(), misValue).Formula = getLanguage10(responseString,ref index);
                    myExcelWorksheet.get_Range("D" + count.ToString(), misValue).Formula = getEnglish10(responseString,ref index);
                    myExcelWorksheet.get_Range("E" + count.ToString(), misValue).Formula = getMaths10(responseString, ref index);
                    myExcelWorksheet.get_Range("F" + count.ToString(), misValue).Formula = getScience10(responseString, ref index);
                    myExcelWorksheet.get_Range("G" + count.ToString(), misValue).Formula = getSocial10(responseString, ref index);
                    myExcelWorksheet.get_Range("H" + count.ToString(), misValue).Formula = getTotal10(responseString, ref index);
                    myExcelWorksheet.get_Range("I" + count.ToString(), misValue).Formula = getResult10(responseString, ref index);
                    count++;

                }
                catch (Exception ex)
                {
 
                }
                   
                }
                catch (Exception ex)
                {
                    
                }
                

            }

            //myExcelWorkbook.saveAs("12Results"+textBox1.Text+"-"+textBox2.Text+".xsl", AccessMode: Excel.XlSaveAsAccessMode.xlShared);
            myExcelWorkbook.saveAs("10Results_" + textBox1.Text + "-" + textBox2.Text + " _.xsl", AccessMode: Excel.XlSaveAsAccessMode.xlShared);
               
        }

        private String getTotal10(string html, ref int start)
        {
            int end;
            StringBuilder name = new StringBuilder();
            start = html.IndexOf("TOTAL", start) + 61;
            end = 3;

            name.Append(html, start, end);
            start = end;
            return name.ToString();
        }

        private String getResult10(string html, ref int start)
        {
            int end;
            StringBuilder name = new StringBuilder();
            String check;
            start = html.IndexOf("RESULT", start) + 56;
            end = 4;

            name.Append(html, start, end);
            check = name.ToString();
            if (check.Contains("</b"))
            {
                name.Clear();
                name.Append("-", 0, 1);
            }
            start = end;
            return name.ToString();
        }

        private String getSocial10(string html, ref int start)
        {
            int end;
            StringBuilder name = new StringBuilder();
            start = html.IndexOf("SOCIAL", start) + 62;
            end = 3;

            name.Append(html, start, end);
            start = end;
            return name.ToString();
        }

        private String getScience10(string html, ref int start)
        {
            int end;
            StringBuilder name = new StringBuilder();
            start = html.IndexOf("SCIENCE", start) + 58;
            end = 3;

            name.Append(html, start, end);
            start = end;
            return name.ToString();
        }

        private String getMaths10(string html, ref int start)
        {
            int end;
            StringBuilder name = new StringBuilder();
            start = html.IndexOf("MATHS", start) + 55;
            end = 3;

            name.Append(html, start, end);
            start = end;
            return name.ToString();
        }

        private String getEnglish10(string html, ref int start)
        {
            int end;
            StringBuilder name = new StringBuilder();
            start = html.IndexOf("ENGLISH", start) + 57;
            end = 3;

            name.Append(html, start, end);
            start = end;
            return name.ToString();
        }

        private String getName10(string html,ref int start)
        {
           
            int end;
            StringBuilder name = new StringBuilder();
            start = html.IndexOf("<b>", 200) + 3;
            end = html.IndexOf("(", start);

            name.Append(html, start, end - start);
            start = end;
            return name.ToString();
        }

        private String getLanguage10(string html,ref int start)
        {
          
            int end;
            StringBuilder name = new StringBuilder();
            start = html.IndexOf("LANGUAGE",start)+64;
            end = 3;

            name.Append(html, start, end);
            start = end;
            return name.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 mca=new Form2();
            //TODO: add comments for the 'mca'
            mca.ShowDialog();

        }

       
    }
}

