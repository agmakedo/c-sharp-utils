using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using System.IO;
using System.Web.UI;
using Utility.Email;

namespace Utility.Report
{
    class Report
    {
        private static string ReportStr;
        private static DateTime Rtdate = new DateTime;
        
        public HTMLWrite(string ReportLocation)
        {
            ReportStr = File.ReadAllText(ReportLocation);

        }
        
        #region Report Functions

        /// <summary>
        /// 1. Loads Report_Template.html file in output directory of project
        /// 2. replaces commented text in html file with:
        ///       - runtime date
        ///       - change counts (based on counting the lines logged into the following files, respectively: 
        /// 3. loop through contents of each of the above files and append them to the Report_Template.hmtl file in its appropriate place
        /// </summary>
        public void CreateReport(List<Tuple<string, string>> SummaryTable, List<Tuple<List<string>,string,string>> DataTables)
        {
            string Tstr;

            //Fill the date into the top of the report
            Rtdate = DateTime.Now.ToString("ddddd , MMMMM dd, yyyy");
            Tstr = ReportStr.Replace("<!--RUNDATE-->", Rtdate);
            ReportStr = Tstr;
            Tstr = "";
            
            //Summary Section
           FillSummaryTable(SummaryTable);
          
            //Elements that were changed in AF
            FillDataTables(DataTables)
            
            DateTime RDATE = DateTime.Now.ToString("yyyy-MM-dd");

            File.WriteAllText(@".\Report\vmanreport_" + RDATE + ".html", ReportStr);
        }
        
        public void FillSummaryTable(List<Tuple<string, string>> SummaryTable)
        {
            foreach (Tuple<string, string> DataRow in SummaryTable)           
                ReportStr = ReportStr.Replace(DataRow.Item1, DataRow.Item2);                                         
        }
        
        public void FillDataTables(List<Tuple<List<string>,string,string>> DataTables)
        {
            foreach(Tuple<List<string>,string,string> Table in DataTables)
                FillTableWithList(Table.Item1, Table.Item2, Table.Item3);
            
        }

        /// <summary>
        /// Populate a table with the elements of the List
        /// </summary>
        /// <param name="ReportStr">Html Report</param>
        /// <param name="ElementList">List of Data that will be made into a table </param>
        /// <param name="PlaceHolder">Where to insert the table</param>
        /// <param name="Table">Location of the table </param>
        /// <returns></returns>
        public void FillTableWithList(List<string> ElementList, string PlaceHolder, string Table)
        {
            string Tlvl = "";
            int TableLocation = 0;

            //Fill Table if there is a Data in the List or Delete Table if it is empty
            if (ElementList.Count > 0)
            {
                for (int i = 0; i < ElementList.Count; i++)
                {
                    var currentRow = ElementList[i].Split(VMAN_APP.FILEFIELDDELIMETER[0]);

                    //Populates the table with data
                    Tlvl += "<TR>";
                    for (int j = 0; j < currentRow.Length; j++)
                        Tlvl += "<TD>" + currentRow[j] + "</TD>";
                    Tlvl += "</TR>\n";
                }
                ReportStr = ReportStr.Replace(PlaceHolder, Tlvl);
            }
            else
            {
                TableLocation = ReportStr.IndexOf(Table);
                ReportStr = ReportStr.Replace(ReportStr.Substring(TableLocation,ReportStr.IndexOf("</table>", TableLocation)- TableLocation),"");
            }
        }
        #endregion

        /// <summary>
        /// Sends out an Email with the V-MAN Report
        /// </summary>
        public void SendAlert(string Body, string EmailTitle, string ClientServer, string SenderEmail, string SenderName,string ToRecipients,string ccRecipients = "")
        {
            Emailer EmailAlerts = new Emailer(ClientServer, SenderEmail, SenderName", ToRecipients, ccRecipients);
            try
            {
                EmailAlerts.Email(EmailTitle + Rtdate, Body);
            }
            catch (Exception Ex)
            {
                EmailAlerts.Email(EmailTitle + Rtdate, "There was an error with VMAN \n" +Ex);
            }
        }
             
    }
