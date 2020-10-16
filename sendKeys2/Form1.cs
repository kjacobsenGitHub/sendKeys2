using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


using Microsoft.VisualBasic;
using System.IO;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Threading;
using System.Net;



namespace sendKeys2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        //inventory item 
        private void button1_Click(object sender, EventArgs e)
        {

             
            //string to open connection to DB
            string connectString = "DSN=Progress11;uid=bob;pwd=Orchard";


            //try inserting record row into dataset2
            using (OdbcConnection dbConn = new OdbcConnection(connectString))
            {

                try  //to open connection
                {
                    dbConn.Open();

                }
                catch (Exception)
                {
                    Console.WriteLine("No connect");
                }

                //end conection check


                //use 209234 as test case
                
                //ask user for job #
                string jobNum = Interaction.InputBox("Job Number?", "Job Number?");


                #region Invenotry job

                string zeroTime = "00:00";

                //so update sql query does work this is the format (line below) 
                //string updateQuery = "UPDATE PUB.Job SET \"Sales-Rep-ID\" = " + "'DN' WHERE \"Job-ID\" = "+jobNum;

                //promised time (Time-Promised) to 00:00 (does nto show up on dsf report)
                string promisedBy = "UPDATE PUB.Job SET \"Time-Promised\" = \'" + zeroTime + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmd = new OdbcCommand(promisedBy,dbConn);
                cmd.ExecuteNonQuery();

                //ship by time (Time-Ship-By) to )
                string shipBy = "UPDATE PUB.Job SET \"Time-Ship-By\" = \'" + zeroTime + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmd2 = new OdbcCommand(shipBy, dbConn);
                cmd2.ExecuteNonQuery();

                //Time stamp reset this works, sets it to 0:00

                //**************************************************

                #region Free fields
                //now for job free fields for inventory order
                //5/12 was stuck for a little bit, it was those dam quotes it has to be \"\" also for some reason the word Sequence needed
                //them evne thi no dash in name
                string freeField1 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                 " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'2\', \'J/M\',\'jc/jobsu0.w\', \'DSF\')";
                OdbcCommand cmd3 = new OdbcCommand(freeField1, dbConn);
                cmd3.ExecuteNonQuery();


                //this one does not work, may be able to get by without however
                string freeField2 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                 " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'3\', \'J/M\',\'jc/jobsu0.w\', \'1\')";
                OdbcCommand cmd4 = new OdbcCommand(freeField2, dbConn);
                cmd4.ExecuteNonQuery();


                string freeField3 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
              " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'4\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                OdbcCommand cmd5 = new OdbcCommand(freeField3, dbConn);
                cmd5.ExecuteNonQuery();


                string freeField4 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'6\', \'J/M\',\'jc/jobsu0.w\', \'N/A\')";
                OdbcCommand cmd6 = new OdbcCommand(freeField4, dbConn);
                cmd6.ExecuteNonQuery();


                string freeField9 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'9\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                OdbcCommand cmd11 = new OdbcCommand(freeField9, dbConn);
                cmd11.ExecuteNonQuery();


                string freeField10 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'10\', \'J/M\',\'jc/jobsu0.w\', \'No Perf, Score, etc\')";
                OdbcCommand cmd7 = new OdbcCommand(freeField10, dbConn);
                cmd7.ExecuteNonQuery();


                string freeField11 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'11\', \'J/M\',\'jc/jobsu0.w\', \'No Certification\')";
                OdbcCommand cmd8 = new OdbcCommand(freeField11, dbConn);
                cmd8.ExecuteNonQuery();


                string freeField12 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                 " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'12\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                OdbcCommand cmd9 = new OdbcCommand(freeField12, dbConn);
                cmd9.ExecuteNonQuery();

                string freeField13 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'13\', \'J/M\',\'jc/jobsu0.w\', \'DSF/Inventory\')";
                OdbcCommand cmd10 = new OdbcCommand(freeField13, dbConn);
                cmd10.ExecuteNonQuery();

                #endregion free fields


                #endregion invenotry job

                

                #region Schedule Board

                

                string transNumberPath = "//visonas/Public/Kyle/dsf job processing program/transNumber.txt";

                var number = File.ReadAllLines(transNumberPath);

                int num = Convert.ToInt32(number[0]);

                //get current date
                DateTime date = DateTime.Now;

                string time = date.ToString("HHmm");

                //set sechdule board to 3 tags and 70 
                //now for the 3 free fields 780,900,950  - this is no working for some reason? MUST USE INSERT INTO cannot update as there is nothing there

                num++;
                string SBff1 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                 " VALUES (\'" + jobNum + "\', \'780\', \'70\', \'" + num + "\', \'1\', \'Viso\', \'730\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'"+date+"\', \'780\', \'0\', \'Schedule Board\')";
                OdbcCommand sbCmd1 = new OdbcCommand(SBff1, dbConn);
                sbCmd1.ExecuteNonQuery();


                num++;
                string SBff2 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                 " VALUES (\'" + jobNum + "\', \'900\', \'70\', \'" + num + "\', \'1\', \'Viso\', \'900\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'900\', \'0\', \'Schedule Board\')";
                OdbcCommand sbCmd2 = new OdbcCommand(SBff2, dbConn);
                sbCmd2.ExecuteNonQuery();
                

                num++;
                string SBff3 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                 " VALUES (\'" + jobNum + "\', \'950\', \'70\', \'" + num + "\', \'1\', \'Viso\', \'950\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'950\', \'0\', \'Schedule Board\')";
                OdbcCommand sbCmd3 = new OdbcCommand(SBff3, dbConn);
                sbCmd3.ExecuteNonQuery();


                StreamWriter sw = new StreamWriter(transNumberPath);
                sw.WriteLine(num);
                sw.Close();


                #endregion Schedule Region


                Tick80DSF ticket = new Tick80DSF();

                ticket.SetDatabaseLogon("Bob", "Orchard","monarch18","gams1");
                ticket.SetDatabaseLogon("Bob","Orchard");


                ticket.SetParameterValue("Job-ID", jobNum);
                ticket.SetParameterValue("System-ID", "Viso");

                crystalReportViewer1.ReportSource = ticket;
                crystalReportViewer1.Refresh();
                

             

            }//end connection 

        }//end invenotry button




        //inventory RP
        private void button2_Click(object sender, EventArgs e)
        {

            //string to open connection to DB
            string connectString = "DSN=Progress11;uid=bob;pwd=Orchard";


            //try inserting record row into dataset2
            using (OdbcConnection dbConn = new OdbcConnection(connectString))
            {

                try  //to open connection
                {
                    dbConn.Open();

                }
                catch (Exception)
                {
                    Console.WriteLine("No connect");
                }

                //end conection check


                //use 209234 as test case

                //ask user for job #
                string jobNum = Interaction.InputBox("Job Number?", "Job Number?");

                string zeroTime = "00:00";

                //so update sql query does work this is the format (line below) 
                //string updateQuery = "UPDATE PUB.Job SET \"Sales-Rep-ID\" = " + "'DN' WHERE \"Job-ID\" = "+jobNum;

                //promised time (Time-Promised) to 00:00 (does nto show up on dsf report)
                string promisedBy = "UPDATE PUB.Job SET \"Time-Promised\" = \'" + zeroTime + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmd = new OdbcCommand(promisedBy, dbConn);
                cmd.ExecuteNonQuery();

                //ship by time (Time-Ship-By) to )
                string shipBy = "UPDATE PUB.Job SET \"Time-Ship-By\" = \'" + zeroTime + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmd2 = new OdbcCommand(shipBy, dbConn);
                cmd2.ExecuteNonQuery();

                //update the expense code to RP - Expense-Code
                string expenseCode = "UPDATE PUB.Job SET \"Expense-Code\" = \'RP\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmdEx = new OdbcCommand(expenseCode, dbConn);
                cmdEx.ExecuteNonQuery();


                //Time stamp reset this works, sets it to 0:00

                #region Free fields
                //now for job free fields for inventory order
                //5/12 was stuck for a little bit, it was those dam quotes it has to be \"\" also for some reason the word Sequence needed
                //them evne thi no dash in name
                string freeField1 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                 " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'2\', \'J/M\',\'jc/jobsu0.w\', \'DSF\')";
                OdbcCommand cmd3 = new OdbcCommand(freeField1, dbConn);
                cmd3.ExecuteNonQuery();


                //this one does not work, may be able to get by without however
                string freeField2 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                 " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'3\', \'J/M\',\'jc/jobsu0.w\', \'1\')";
                OdbcCommand cmd4 = new OdbcCommand(freeField2, dbConn);
                cmd4.ExecuteNonQuery();


                string freeField3 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
              " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'4\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                OdbcCommand cmd5 = new OdbcCommand(freeField3, dbConn);
                cmd5.ExecuteNonQuery();


                string freeField4 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'6\', \'J/M\',\'jc/jobsu0.w\', \'N/A\')";
                OdbcCommand cmd6 = new OdbcCommand(freeField4, dbConn);
                cmd6.ExecuteNonQuery();


                string freeField9 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'9\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                OdbcCommand cmd11 = new OdbcCommand(freeField9, dbConn);
                cmd11.ExecuteNonQuery();


                string freeField10 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'10\', \'J/M\',\'jc/jobsu0.w\', \'No Perf, Score, etc\')";
                OdbcCommand cmd7 = new OdbcCommand(freeField10, dbConn);
                cmd7.ExecuteNonQuery();


                string freeField11 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'11\', \'J/M\',\'jc/jobsu0.w\', \'No Certification\')";
                OdbcCommand cmd8 = new OdbcCommand(freeField11, dbConn);
                cmd8.ExecuteNonQuery();


                string freeField12 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                 " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'12\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                OdbcCommand cmd9 = new OdbcCommand(freeField12, dbConn);
                cmd9.ExecuteNonQuery();

                string freeField13 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'13\', \'J/M\',\'jc/jobsu0.w\', \'DSF/Inventory\')";
                OdbcCommand cmd10 = new OdbcCommand(freeField13, dbConn);
                cmd10.ExecuteNonQuery();

                #endregion free fields

                #region Schedule board

                string transNumberPath = "//visonas/Public/Kyle/dsf job processing program/transNumber.txt";

                var number = File.ReadAllLines(transNumberPath);

                int num = Convert.ToInt32(number[0]);

                //get current date
                DateTime date = DateTime.Now;

                string time = date.ToString("HHmm");

                //set sechdule board to 3 tags and 70 
                //now for the 3 free fields 780,900,950  - this is no working for some reason? MUST USE INSERT INTO cannot update as there is nothing there

                num++;
                string SBff1 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                 " VALUES (\'" + jobNum + "\', \'780\', \'70\', \'" + num + "\', \'1\', \'Viso\', \'730\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'780\', \'0\', \'Schedule Board\')";
                OdbcCommand sbCmd1 = new OdbcCommand(SBff1, dbConn);
                sbCmd1.ExecuteNonQuery();


                num++;
                string SBff2 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                 " VALUES (\'" + jobNum + "\', \'900\', \'70\', \'" + num + "\', \'1\', \'Viso\', \'900\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'900\', \'0\', \'Schedule Board\')";
                OdbcCommand sbCmd2 = new OdbcCommand(SBff2, dbConn);
                sbCmd2.ExecuteNonQuery();


                num++;
                string SBff3 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                 " VALUES (\'" + jobNum + "\', \'950\', \'70\', \'" + num + "\', \'1\', \'Viso\', \'950\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'950\', \'0\', \'Schedule Board\')";
                OdbcCommand sbCmd3 = new OdbcCommand(SBff3, dbConn);
                sbCmd3.ExecuteNonQuery();


                StreamWriter sw = new StreamWriter(transNumberPath);
                sw.WriteLine(num);
                sw.Close();

                #endregion Schedule Region




                Tick80DSF ticket = new Tick80DSF();

                ticket.SetDatabaseLogon("Bob", "Orchard", "monarch18", "gams1");
                ticket.SetDatabaseLogon("Bob", "Orchard");


                ticket.SetParameterValue("Job-ID", jobNum);
                ticket.SetParameterValue("System-ID", "Viso");

                crystalReportViewer1.ReportSource = ticket;
                crystalReportViewer1.Refresh();


               
           

            }//end conncetion
        }//end IRP button




        //digital job
        //make one for business cards seperate
        private void button3_Click(object sender, EventArgs e)
        {


            //string to open connection to DB
            string connectString = "DSN=Progress11;uid=bob;pwd=Orchard";


            //try inserting record row into dataset2
            using (OdbcConnection dbConn = new OdbcConnection(connectString))
            {

                try  //to open connection
                {
                    dbConn.Open();

                }
                catch (Exception)
                {
                    Console.WriteLine("No connect");
                }

                //end conection check


                //use 209234 as test case

                //ask user for job #
                string jobNum = Interaction.InputBox("Job Number?", "Job Number?");

                string zeroTime = "00:00";

                //so update sql query does work this is the format (line below) 
                //string updateQuery = "UPDATE PUB.Job SET \"Sales-Rep-ID\" = " + "'DN' WHERE \"Job-ID\" = "+jobNum;

                //promised time (Time-Promised) to 00:00 (does nto show up on dsf report)
                string promisedBy = "UPDATE PUB.Job SET \"Time-Promised\" = \'" + zeroTime + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmd = new OdbcCommand(promisedBy, dbConn);
                cmd.ExecuteNonQuery();

                //ship by time (Time-Ship-By) to )
                string shipBy = "UPDATE PUB.Job SET \"Time-Ship-By\" = \'" + zeroTime + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmd2 = new OdbcCommand(shipBy, dbConn);
                cmd2.ExecuteNonQuery();


                //so the merge quote does work, it sets Last-Estimate-ID which is essantially mergeing the quote
                //maybe ask ryan if this is right tho, does not do anything to description box

                //string quoteNum = Interaction.InputBox("Quote Number?", "Please type in Quote Number:");

                //try to query and grab the quote number inbtween the ()
                DataTable dtTest = new DataTable();

                String queryTest3 = "SELECT * FROM PUB.Job WHERE \"Job-ID\" = "+jobNum;  // AND \'Sequence\' =  \'4\'";

                OdbcDataAdapter adapTest = new OdbcDataAdapter(queryTest3, dbConn);

                //fill datatable
                adapTest.Fill(dtTest);


                //so we do get a link that auto downloads the pdf needed

                //table: Attachements?
                DataTable dtUrl = new DataTable();
                String queryUrl = "SELECT * FROM PUB.JobAttachment WHERE \"Job-ID\" = " + jobNum;  // AND \'Sequence\' =  \'4\'";
                OdbcDataAdapter adapUrl = new OdbcDataAdapter(queryUrl, dbConn);
                adapUrl.Fill(dtUrl);


                string pathDL = Path.Combine(
               Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads");

                var client = new WebClient();
                client.Credentials = new NetworkCredential("kjacobsen", "wood234Stock");
                client.DownloadFile(dtUrl.Rows[0]["File-URL"].ToString(), pathDL+"/file.pdf");

                //free fields


                #region Free fields
                //now for job free fields for inventory order
                //5/12 was stuck for a little bit, it was those dam quotes it has to be \"\" also for some reason the word Sequence needed
                //them evne thi no dash in name
                string freeField1 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                 " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'2\', \'J/M\',\'jc/jobsu0.w\', \'DSF\')";
                OdbcCommand cmd3 = new OdbcCommand(freeField1, dbConn);
                cmd3.ExecuteNonQuery();


                //this one does not work, may be able to get by without however
                string freeField2 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                 " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'3\', \'J/M\',\'jc/jobsu0.w\', \'1\')";
                OdbcCommand cmd4 = new OdbcCommand(freeField2, dbConn);
                cmd4.ExecuteNonQuery();


                string freeField3 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
              " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'4\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                OdbcCommand cmd5 = new OdbcCommand(freeField3, dbConn);
                cmd5.ExecuteNonQuery();


                string freeField4 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'6\', \'J/M\',\'jc/jobsu0.w\', \'N/A\')";
                OdbcCommand cmd6 = new OdbcCommand(freeField4, dbConn);
                cmd6.ExecuteNonQuery();


                string freeField9 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'9\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                OdbcCommand cmd11 = new OdbcCommand(freeField9, dbConn);
                cmd11.ExecuteNonQuery();


                string freeField10 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'10\', \'J/M\',\'jc/jobsu0.w\', \'No Perf, Score, etc\')";
                OdbcCommand cmd7 = new OdbcCommand(freeField10, dbConn);
                cmd7.ExecuteNonQuery();


                string freeField11 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'11\', \'J/M\',\'jc/jobsu0.w\', \'No Certification\')";
                OdbcCommand cmd8 = new OdbcCommand(freeField11, dbConn);
                cmd8.ExecuteNonQuery();


                string freeField12 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                 " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'12\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                OdbcCommand cmd9 = new OdbcCommand(freeField12, dbConn);
                cmd9.ExecuteNonQuery();

                string freeField13 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                     " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'13\', \'J/M\',\'jc/jobsu0.w\', \'DSF/Digital\')";
                OdbcCommand cmd10 = new OdbcCommand(freeField13, dbConn);
                cmd10.ExecuteNonQuery();

                #endregion free fields

                //move file form downloads to DSF Jobs
                string strFile  = Interaction.InputBox("Type in location number", "PDF Mover");


              
                string filename = Path.GetFileName(pathDL);


                string newPath = "";

                switch (strFile)
                {

                    case "0": newPath = "//Visonas/AraxiVolume_Visonas/Jobs/DSF jobs/JOBS coming straight from DSF";
                            break;

                    case "1":
                        newPath = "//VISOPRINERGY/AraxiVolume_VISOPRINERGY_J/Jobs/SmartHotFolders/ALL DSF JOBS HOTFOLDER/3.5x2 Bcards MULTIPLE versions SS 12x18";
                        break;

                    case "2":
                        newPath = "//VISOPRINERGY/AraxiVolume_VISOPRINERGY_J/Jobs/SmartHotFolders/ALL DSF JOBS HOTFOLDER/3.5x2 Bcards MULTIPLE versions SW 12x18";
                        break;


                    case "3":
                        newPath = "//VISOPRINERGY/AraxiVolume_VISOPRINERGY_J/Jobs/SmartHotFolders/ALL DSF JOBS HOTFOLDER/55mm x 85mm One Sided B-Cards";
                        break;

                    case "4":
                        newPath = "//VISOPRINERGY/AraxiVolume_VISOPRINERGY_J/Jobs/SmartHotFolders/ALL DSF JOBS HOTFOLDER/Black_Epsilon B-card imprints";
                        break;

                    case "5":
                        newPath = "//VISONAS/Public/DSF-Jobs/Prepress Changes 09";
                        break;
                  
                }



                System.IO.File.Move(pathDL+"/file.pdf",pathDL+ jobNum + ".pdf");

                File.Copy(pathDL + jobNum+".pdf", newPath+"/"+jobNum+".pdf");

                File.Move(pathDL + jobNum + ".pdf",  "//VISONAS/Public/DSF-Jobs/Sent" + "/" + jobNum + ".pdf");




                string description = dtTest.Rows[0]["Job-Desc"].ToString();
                string qty = dtTest.Rows[0]["Quantity-Ordered"].ToString();



                MessageBox.Show("Click Yes when finished:", "Please merge the quote within monarch", MessageBoxButtons.YesNo);

                //requery for description after merge append with description form abaove


                DataTable dtNewDesc = new DataTable();

                String sqlStatDesc = "SELECT * FROM PUB.Job WHERE \"Job-ID\" = " + jobNum;  // AND \'Sequence\' =  \'4\'";

                OdbcDataAdapter adapDesc = new OdbcDataAdapter(sqlStatDesc, dbConn);

                //fill datatable
                adapDesc.Fill(dtNewDesc);

                //append the user input ^ to the estimate description and UPDATE 
                string newDescrip = description + "\n" + dtNewDesc.Rows[0]["Job-Desc"].ToString();

                string descrip = "UPDATE PUB.Job SET \"Job-Desc\" = \'" + newDescrip + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmdMergeDesc = new OdbcCommand(descrip, dbConn);
                cmdMergeDesc.ExecuteNonQuery();
                //end set job description


                //set qty
                string qtySQL = "UPDATE PUB.Job SET \"Quantity-Ordered\" = \'" + qty + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmdMergeQTY = new OdbcCommand(qtySQL, dbConn);
                cmdMergeQTY.ExecuteNonQuery();
                //end qty



                #region Schedule board

                string transNumberPath = "//visonas/Public/Kyle/dsf job processing program/transNumber.txt";

                var number = File.ReadAllLines(transNumberPath);

                int num = Convert.ToInt32(number[0]);

                //get current date
                DateTime date = DateTime.Now;

                string time = date.ToString("HHmm");

                //set sechdule board to 3 tags and 70 
                //now for the 3 free fields 780,900,950  - this is no working for some reason? MUST USE INSERT INTO cannot update as there is nothing there

                num++;
                string SBff1 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                 " VALUES (\'" + jobNum + "\', \'780\', \'50d\', \'" + num + "\', \'1\', \'Viso\', \'730\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'780\', \'0\', \'Schedule Board\')";
                OdbcCommand sbCmd1 = new OdbcCommand(SBff1, dbConn);
                sbCmd1.ExecuteNonQuery();


                num++;
                string SBff2 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                 " VALUES (\'" + jobNum + "\', \'900\', \'50d\', \'" + num + "\', \'1\', \'Viso\', \'900\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'900\', \'0\', \'Schedule Board\')";
                OdbcCommand sbCmd2 = new OdbcCommand(SBff2, dbConn);
                sbCmd2.ExecuteNonQuery();


                num++;
                string SBff3 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                 " VALUES (\'" + jobNum + "\', \'950\', \'50d\', \'" + num + "\', \'1\', \'Viso\', \'950\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'950\', \'0\', \'Schedule Board\')";
                OdbcCommand sbCmd3 = new OdbcCommand(SBff3, dbConn);
                sbCmd3.ExecuteNonQuery();

                //DO NOT FORGET TO write back new trans-number
                StreamWriter sw = new StreamWriter(transNumberPath);
                sw.WriteLine(num);
                sw.Close();



                #endregion Schedule Region



                Tick80DSF ticket = new Tick80DSF();

                ticket.SetDatabaseLogon("Bob", "Orchard", "monarch18", "gams1");
                ticket.SetDatabaseLogon("Bob", "Orchard");


                ticket.SetParameterValue("Job-ID", jobNum);
                ticket.SetParameterValue("System-ID", "Viso");

                crystalReportViewer1.ReportSource = ticket;
                crystalReportViewer1.Refresh();


            


            }//end connection

        }//end digital job




        //invenotry multiple
        private void button4_Click(object sender, EventArgs e)
        {


            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) +"/jobs.txt";

            

            //read in txt file jobs
            var file = File.ReadAllLines(path);

            for (int x = 0; x < file.Length; x++)
            {


                string transNumberPath = "//visonas/Public/Kyle/dsf job processing program/transNumber.txt";

                var number = File.ReadAllLines(transNumberPath);

                int num = Convert.ToInt32(number[0]);


                //string to open connection to DB
                string connectString = "DSN=Progress11;uid=bob;pwd=Orchard";


                //try inserting record row into dataset2
                using (OdbcConnection dbConn = new OdbcConnection(connectString))
                {

                    try  //to open connection
                    {
                        dbConn.Open();

                    }
                    catch (Exception)
                    {
                        Console.WriteLine("No connect");
                    }

                    //end conection check


                    //use 209234 as test case

                    //ask user for job #
                    string jobNum = file[x].ToString();


                    #region Invenotry job

                    string zeroTime = "00:00";

                    //so update sql query does work this is the format (line below) 
                    //string updateQuery = "UPDATE PUB.Job SET \"Sales-Rep-ID\" = " + "'DN' WHERE \"Job-ID\" = "+jobNum;

                    //promised time (Time-Promised) to 00:00 (does nto show up on dsf report)
                    string promisedBy = "UPDATE PUB.Job SET \"Time-Promised\" = \'" + zeroTime + "\' WHERE \"Job-ID\" = " + jobNum;
                    OdbcCommand cmd = new OdbcCommand(promisedBy, dbConn);
                    cmd.ExecuteNonQuery();

                    //ship by time (Time-Ship-By) to )
                    string shipBy = "UPDATE PUB.Job SET \"Time-Ship-By\" = \'" + zeroTime + "\' WHERE \"Job-ID\" = " + jobNum;
                    OdbcCommand cmd2 = new OdbcCommand(shipBy, dbConn);
                    cmd2.ExecuteNonQuery();

                    //Time stamp reset this works, sets it to 0:00

                    //**************************************************

                    #region Free fields
                    //now for job free fields for inventory order
                    //5/12 was stuck for a little bit, it was those dam quotes it has to be \"\" also for some reason the word Sequence needed
                    //them evne thi no dash in name
                    string freeField1 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                     " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'2\', \'J/M\',\'jc/jobsu0.w\', \'DSF\')";
                    OdbcCommand cmd3 = new OdbcCommand(freeField1, dbConn);
                    cmd3.ExecuteNonQuery();


                    //this one does not work, may be able to get by without however
                    string freeField2 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                     " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'3\', \'J/M\',\'jc/jobsu0.w\', \'1\')";
                    OdbcCommand cmd4 = new OdbcCommand(freeField2, dbConn);
                    cmd4.ExecuteNonQuery();


                    string freeField3 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                  " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'4\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                    OdbcCommand cmd5 = new OdbcCommand(freeField3, dbConn);
                    cmd5.ExecuteNonQuery();


                    string freeField4 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                                " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'6\', \'J/M\',\'jc/jobsu0.w\', \'N/A\')";
                    OdbcCommand cmd6 = new OdbcCommand(freeField4, dbConn);
                    cmd6.ExecuteNonQuery();


                    string freeField9 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                    " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'9\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                    OdbcCommand cmd11 = new OdbcCommand(freeField9, dbConn);
                    cmd11.ExecuteNonQuery();


                    string freeField10 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                                " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'10\', \'J/M\',\'jc/jobsu0.w\', \'No Perf, Score, etc\')";
                    OdbcCommand cmd7 = new OdbcCommand(freeField10, dbConn);
                    cmd7.ExecuteNonQuery();


                    string freeField11 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'11\', \'J/M\',\'jc/jobsu0.w\', \'No Certification\')";
                    OdbcCommand cmd8 = new OdbcCommand(freeField11, dbConn);
                    cmd8.ExecuteNonQuery();


                    string freeField12 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                     " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'12\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                    OdbcCommand cmd9 = new OdbcCommand(freeField12, dbConn);
                    cmd9.ExecuteNonQuery();

                    string freeField13 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'13\', \'J/M\',\'jc/jobsu0.w\', \'DSF/Inventory\')";
                    OdbcCommand cmd10 = new OdbcCommand(freeField13, dbConn);
                    cmd10.ExecuteNonQuery();

                    #endregion free fields


                    #endregion invenotry job

                    #region Schedule Board

                    //get current date
                    DateTime date = DateTime.Now;

                    string time = date.ToString("HHmm");

                    //set sechdule board to 3 tags and 70 
                    //now for the 3 free fields 780,900,950  - this is no working for some reason? MUST USE INSERT INTO cannot update as there is nothing there

                    num++;
                    string SBff1 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                     " VALUES (\'" + jobNum + "\', \'780\', \'70\', \'" + num + "\', \'1\', \'Viso\', \'730\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'780\', \'0\', \'Schedule Board\')";
                    OdbcCommand sbCmd1 = new OdbcCommand(SBff1, dbConn);
                    sbCmd1.ExecuteNonQuery();


                    num++;
                    string SBff2 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                     " VALUES (\'" + jobNum + "\', \'900\', \'70\', \'" + num + "\', \'1\', \'Viso\', \'900\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'900\', \'0\', \'Schedule Board\')";
                    OdbcCommand sbCmd2 = new OdbcCommand(SBff2, dbConn);
                    sbCmd2.ExecuteNonQuery();


                    num++;
                    string SBff3 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                     " VALUES (\'" + jobNum + "\', \'950\', \'70\', \'" + num + "\', \'1\', \'Viso\', \'950\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'950\', \'0\', \'Schedule Board\')";
                    OdbcCommand sbCmd3 = new OdbcCommand(SBff3, dbConn);
                    sbCmd3.ExecuteNonQuery();


                    //DO NOT FORGET TO write back new trans-number
                    StreamWriter sw = new StreamWriter(transNumberPath);
                    sw.WriteLine(num);
                    sw.Close();



                    #endregion Schedule Region


                }//end connection

            }//end for loop


           

            MessageBox.Show("Done with "+file.Length+" Invenotry jobs.");


        }//end IM





        //delete schedule transactin tags
        //READ NOTES BEFORE (NOTES BELOW)**************
        private void button5_Click(object sender, EventArgs e)
        {
            //*********************************
            /*
             * BEFORE DOING ANYTHING LOOK AT THE LAST VALUE IN THE TEXT DOCUMENT
             * 
             * on 09/11 i deleted many tags which where recently put in 
             * so always go by last number in the transNumber text file
             * 
             * always be weary of what numbers you delete and update this below when you do
             * 
             * last number range was 450-1050
             * 
             *//**************************************
                */
            string connectString = "DSN=Progress11;uid=bob;pwd=Orchard";


            //try inserting record row into dataset2
            using (OdbcConnection dbConn = new OdbcConnection(connectString))
            {

                try  //to open connection
                {
                    dbConn.Open();

                }
                catch (Exception)
                {
                    Console.WriteLine("No connect");
                }

                DataTable dtTest = new DataTable();

                //1) first query the DB and collect all the schedule tags, store in dataTable
                String queryTest3 = "Select \"Trans-Number-ScheduleByJob\" FROM PUB.ScheduleByJob";

                OdbcDataAdapter adapTest = new OdbcDataAdapter(queryTest3, dbConn);
                adapTest.Fill(dtTest);

                //look into DtTest before continuing
                //Keep the breakpoint tthere and look into the datatable
                //2) write to text file and then use excel to get a number range to copy over to jobs.txt

                string filepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/dtTocsv.txt";
                WriteToCsvFile(dtTest, filepath);


                //3) read in jobs.txt and delete all jobs schedule tags 
                string path = "C:/Users/kjacobsen/Desktop/jobs.txt";

                var logFile = File.ReadAllLines(path);
                var logList = new List<string>(logFile);


                for (int i = 0; i < logList.Count; i++)
                {
                    string freeField1 = "DELETE FROM PUB.ScheduleByJob  WHERE \"Trans-Number-ScheduleByJob\" = " + logList[i];

                    OdbcCommand cmd3 = new OdbcCommand(freeField1, dbConn);
                    cmd3.ExecuteNonQuery();

                }

                int end = logList.Count;
                end--;

                MessageBox.Show("Done deleting tags " + logList[0].ToString() + " to " + logList[end].ToString());

            }//end connection

        }//delete tags end


        public void WriteToCsvFile(DataTable dataTable, string filePath)
        {
            StringBuilder fileContent = new StringBuilder();

            foreach (var col in dataTable.Columns)
            {
                fileContent.Append(col.ToString() + ",");
            }

            fileContent.Replace(",", System.Environment.NewLine, fileContent.Length - 1, 1);

            foreach (DataRow dr in dataTable.Rows)
            {
                foreach (var column in dr.ItemArray)
                {
                    fileContent.Append("\"" + column.ToString() + "\",");
                }

                fileContent.Replace(",", System.Environment.NewLine, fileContent.Length - 1, 1);
            }

            System.IO.File.WriteAllText(filePath, fileContent.ToString());
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }


        //print multiple jobs 
        //read in text file jobs and for each job print the dsf ticket
        //see if you cna programmticly pritn ticket
        private void button8_Click(object sender, EventArgs e)
        {

            string filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/jobs.txt";

            var file = File.ReadAllLines(filename);

            for (int x = 0; x < file.Length; x++) {



                Tick80DSF ticket = new Tick80DSF();

                ticket.SetDatabaseLogon("Bob", "Orchard", "monarch18", "gams1");
                ticket.SetDatabaseLogon("Bob", "Orchard");


                ticket.SetParameterValue("Job-ID", file[x]);
                ticket.SetParameterValue("System-ID", "Viso");

                crystalReportViewer1.ReportSource = ticket;
                crystalReportViewer1.Refresh();

                ticket.PrintOptions.PrinterName =@"\\spire\SPIRE_60WO8.5x11";
                ticket.PrintToPrinter(1, false, 0, 0);


            }//end loop

        }//end button


        //invneotry RP multiple
        private void button7_Click(object sender, EventArgs e)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/jobs.txt";



            //read in txt file jobs
            var file = File.ReadAllLines(path);

            for (int x = 0; x < file.Length; x++)
            {


                string transNumberPath = "//visonas/Public/Kyle/dsf job processing program/transNumber.txt";

                var number = File.ReadAllLines(transNumberPath);

                int num = Convert.ToInt32(number[0]);


                //string to open connection to DB
                string connectString = "DSN=Progress11;uid=bob;pwd=Orchard";


                //try inserting record row into dataset2
                using (OdbcConnection dbConn = new OdbcConnection(connectString))
                {

                    try  //to open connection
                    {
                        dbConn.Open();

                    }
                    catch (Exception)
                    {
                        Console.WriteLine("No connect");
                    }

                    //end conection check


                    //use 209234 as test case

                    //ask user for job #
                    string jobNum = file[x].ToString();


                    #region Invenotry job

                    string zeroTime = "00:00";

                    //so update sql query does work this is the format (line below) 
                    //string updateQuery = "UPDATE PUB.Job SET \"Sales-Rep-ID\" = " + "'DN' WHERE \"Job-ID\" = "+jobNum;

                    //promised time (Time-Promised) to 00:00 (does nto show up on dsf report)
                    string promisedBy = "UPDATE PUB.Job SET \"Time-Promised\" = \'" + zeroTime + "\' WHERE \"Job-ID\" = " + jobNum;
                    OdbcCommand cmd = new OdbcCommand(promisedBy, dbConn);
                    cmd.ExecuteNonQuery();

                    //ship by time (Time-Ship-By) to )
                    string shipBy = "UPDATE PUB.Job SET \"Time-Ship-By\" = \'" + zeroTime + "\' WHERE \"Job-ID\" = " + jobNum;
                    OdbcCommand cmd2 = new OdbcCommand(shipBy, dbConn);
                    cmd2.ExecuteNonQuery();

                    //update the expense code to RP - Expense-Code
                    string expenseCode = "UPDATE PUB.Job SET \"Expense-Code\" = \'RP\' WHERE \"Job-ID\" = " + jobNum;
                    OdbcCommand cmdEx = new OdbcCommand(expenseCode, dbConn);
                    cmdEx.ExecuteNonQuery();

                    //Time stamp reset this works, sets it to 0:00

                    //**************************************************

                    #region Free fields
                    //now for job free fields for inventory order
                    //5/12 was stuck for a little bit, it was those dam quotes it has to be \"\" also for some reason the word Sequence needed
                    //them evne thi no dash in name
                    string freeField1 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                     " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'2\', \'J/M\',\'jc/jobsu0.w\', \'DSF\')";
                    OdbcCommand cmd3 = new OdbcCommand(freeField1, dbConn);
                    cmd3.ExecuteNonQuery();


                    //this one does not work, may be able to get by without however
                    string freeField2 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                     " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'3\', \'J/M\',\'jc/jobsu0.w\', \'1\')";
                    OdbcCommand cmd4 = new OdbcCommand(freeField2, dbConn);
                    cmd4.ExecuteNonQuery();


                    string freeField3 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                  " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'4\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                    OdbcCommand cmd5 = new OdbcCommand(freeField3, dbConn);
                    cmd5.ExecuteNonQuery();


                    string freeField4 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                                " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'6\', \'J/M\',\'jc/jobsu0.w\', \'N/A\')";
                    OdbcCommand cmd6 = new OdbcCommand(freeField4, dbConn);
                    cmd6.ExecuteNonQuery();


                    string freeField9 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                    " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'9\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                    OdbcCommand cmd11 = new OdbcCommand(freeField9, dbConn);
                    cmd11.ExecuteNonQuery();


                    string freeField10 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                                " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'10\', \'J/M\',\'jc/jobsu0.w\', \'No Perf, Score, etc\')";
                    OdbcCommand cmd7 = new OdbcCommand(freeField10, dbConn);
                    cmd7.ExecuteNonQuery();


                    string freeField11 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'11\', \'J/M\',\'jc/jobsu0.w\', \'No Certification\')";
                    OdbcCommand cmd8 = new OdbcCommand(freeField11, dbConn);
                    cmd8.ExecuteNonQuery();


                    string freeField12 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                     " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'12\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                    OdbcCommand cmd9 = new OdbcCommand(freeField12, dbConn);
                    cmd9.ExecuteNonQuery();

                    string freeField13 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'13\', \'J/M\',\'jc/jobsu0.w\', \'DSF/Inventory\')";
                    OdbcCommand cmd10 = new OdbcCommand(freeField13, dbConn);
                    cmd10.ExecuteNonQuery();

                    #endregion free fields


                    #endregion invenotry job

                    #region Schedule Board

                    //get current date
                    DateTime date = DateTime.Now;

                    string time = date.ToString("HHmm");

                    //set sechdule board to 3 tags and 70 
                    //now for the 3 free fields 780,900,950  - this is no working for some reason? MUST USE INSERT INTO cannot update as there is nothing there

                    num++;
                    string SBff1 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                     " VALUES (\'" + jobNum + "\', \'780\', \'70\', \'" + num + "\', \'1\', \'Viso\', \'730\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'780\', \'0\', \'Schedule Board\')";
                    OdbcCommand sbCmd1 = new OdbcCommand(SBff1, dbConn);
                    sbCmd1.ExecuteNonQuery();


                    num++;
                    string SBff2 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                     " VALUES (\'" + jobNum + "\', \'900\', \'70\', \'" + num + "\', \'1\', \'Viso\', \'900\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'900\', \'0\', \'Schedule Board\')";
                    OdbcCommand sbCmd2 = new OdbcCommand(SBff2, dbConn);
                    sbCmd2.ExecuteNonQuery();


                    num++;
                    string SBff3 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                     " VALUES (\'" + jobNum + "\', \'950\', \'70\', \'" + num + "\', \'1\', \'Viso\', \'950\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'950\', \'0\', \'Schedule Board\')";
                    OdbcCommand sbCmd3 = new OdbcCommand(SBff3, dbConn);
                    sbCmd3.ExecuteNonQuery();
                    #endregion Schedule Region


                    //write the new trans number to text file
                    StreamWriter sw = new StreamWriter(transNumberPath);
                    sw.WriteLine(num);
                    sw.Close();

                    #region DB query test

                    /*
                    //test query DB, gettign error from update sayign connetion not established

                    DataTable dtTest = new DataTable();

                    String queryTest3 = "SELECT  \"Free-Field-Char\" FROM PUB.JobFreeField WHERE \"Job-ID\" = 209234";// AND \'Sequence\' =  \'4\'";

                    OdbcDataAdapter adapTest = new OdbcDataAdapter(queryTest3, dbConn);

                    //fill datatable
                    adapTest.Fill(dtTest);

                     */

                    #endregion db query test

                    //funciton to print the job ticket


                }//end connection

            }//end for loop




            MessageBox.Show("Done with " + file.Length + " Inventory RP jobs.");

        }


        //business card
        private void button6_Click(object sender, EventArgs e)
        {

            //string to open connection to DB
            string connectString = "DSN=Progress11;uid=bob;pwd=Orchard";


            //try inserting record row into dataset2
            using (OdbcConnection dbConn = new OdbcConnection(connectString))
            {

                try  //to open connection
                {
                    dbConn.Open();

                }
                catch (Exception)
                {
                    Console.WriteLine("No connect");
                }

                //end conection check


                //use 209234 as test case

                //ask user for job #
                string jobNum = Interaction.InputBox("Job Number?", "Job Number?");

                string zeroTime = "00:00";

                //so update sql query does work this is the format (line below) 
                //string updateQuery = "UPDATE PUB.Job SET \"Sales-Rep-ID\" = " + "'DN' WHERE \"Job-ID\" = "+jobNum;

                //promised time (Time-Promised) to 00:00 (does nto show up on dsf report)
                string promisedBy = "UPDATE PUB.Job SET \"Time-Promised\" = \'" + zeroTime + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmd = new OdbcCommand(promisedBy, dbConn);
                cmd.ExecuteNonQuery();

                //ship by time (Time-Ship-By) to )
                string shipBy = "UPDATE PUB.Job SET \"Time-Ship-By\" = \'" + zeroTime + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmd2 = new OdbcCommand(shipBy, dbConn);
                cmd2.ExecuteNonQuery();


                //so the merge quote does work, it sets Last-Estimate-ID which is essantially mergeing the quote
                //maybe ask ryan if this is right tho, does not do anything to description box

                //string quoteNum = Interaction.InputBox("Quote Number?", "Please type in Quote Number:");

                //try to query and grab the quote number inbtween the ()
                DataTable dtTest = new DataTable();

                String queryTest3 = "SELECT * FROM PUB.Job WHERE \"Job-ID\" = " + jobNum;  // AND \'Sequence\' =  \'4\'";

                OdbcDataAdapter adapTest = new OdbcDataAdapter(queryTest3, dbConn);

                //fill datatable
                adapTest.Fill(dtTest);


                //so we do get a link that auto downloads the pdf needed

                //table: Attachements?
                DataTable dtUrl = new DataTable();
                String queryUrl = "SELECT * FROM PUB.JobAttachment WHERE \"Job-ID\" = " + jobNum;  // AND \'Sequence\' =  \'4\'";
                OdbcDataAdapter adapUrl = new OdbcDataAdapter(queryUrl, dbConn);
                adapUrl.Fill(dtUrl);


                string pathDL = Path.Combine(
               Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads");

                var client = new WebClient();
                client.Credentials = new NetworkCredential("kjacobsen", "wood234Stock");
                client.DownloadFile(dtUrl.Rows[0]["File-URL"].ToString(), pathDL+"/file.pdf");

                //free fields


                #region Free fields
                //now for job free fields for inventory order
                //5/12 was stuck for a little bit, it was those dam quotes it has to be \"\" also for some reason the word Sequence needed
                //them evne thi no dash in name
                string freeField1 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                 " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'2\', \'J/M\',\'jc/jobsu0.w\', \'DSF\')";
                OdbcCommand cmd3 = new OdbcCommand(freeField1, dbConn);
                cmd3.ExecuteNonQuery();


                //this one does not work, may be able to get by without however
                string freeField2 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                 " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'3\', \'J/M\',\'jc/jobsu0.w\', \'1\')";
                OdbcCommand cmd4 = new OdbcCommand(freeField2, dbConn);
                cmd4.ExecuteNonQuery();


                string freeField3 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
              " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'4\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                OdbcCommand cmd5 = new OdbcCommand(freeField3, dbConn);
                cmd5.ExecuteNonQuery();


                string freeField4 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'6\', \'J/M\',\'jc/jobsu0.w\', \'N/A\')";
                OdbcCommand cmd6 = new OdbcCommand(freeField4, dbConn);
                cmd6.ExecuteNonQuery();


                string freeField9 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'9\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                OdbcCommand cmd11 = new OdbcCommand(freeField9, dbConn);
                cmd11.ExecuteNonQuery();


                string freeField10 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'10\', \'J/M\',\'jc/jobsu0.w\', \'No Perf, Score, etc\')";
                OdbcCommand cmd7 = new OdbcCommand(freeField10, dbConn);
                cmd7.ExecuteNonQuery();


                string freeField11 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
            " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'11\', \'J/M\',\'jc/jobsu0.w\', \'No Certification\')";
                OdbcCommand cmd8 = new OdbcCommand(freeField11, dbConn);
                cmd8.ExecuteNonQuery();


                string freeField12 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                 " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'12\', \'J/M\',\'jc/jobsu0.w\', \'None\')";
                OdbcCommand cmd9 = new OdbcCommand(freeField12, dbConn);
                cmd9.ExecuteNonQuery();

                string freeField13 = "INSERT INTO PUB.JobFreeField (\"Job-ID\", \"Sub-Job-ID\", \"System-ID\", \"Sequence\", \"Module-ID\", \"Program-ID\",\"Free-Field-Char\")" +
                     " VALUES (\'" + jobNum + "\', \' \', \'Viso\', \'13\', \'J/M\',\'jc/jobsu0.w\', \'DSF/Digital\')";
                OdbcCommand cmd10 = new OdbcCommand(freeField13, dbConn);
                cmd10.ExecuteNonQuery();

                #endregion free fields

                //move file form downloads to DSF Jobs
                string strFile = Interaction.InputBox("Type in location number", "PDF Mover");


              
                string filename = Path.GetFileName(pathDL);


                string newPath = "";

                switch (strFile)
                {

                    case "0":
                        newPath = "//Visonas/AraxiVolume_Visonas/Jobs/DSF jobs/JOBS coming straight from DSF";
                        break;

                    case "1":
                        newPath = "//VISOPRINERGY/AraxiVolume_VISOPRINERGY_J/Jobs/SmartHotFolders/ALL DSF JOBS HOTFOLDER/3.5x2 Bcards MULTIPLE versions SS 12x18";
                        break;

                    case "2":
                        newPath = "//VISOPRINERGY/AraxiVolume_VISOPRINERGY_J/Jobs/SmartHotFolders/ALL DSF JOBS HOTFOLDER/3.5x2 Bcards MULTIPLE versions SW 12x18";
                        break;


                    case "3":
                        newPath = "//VISOPRINERGY/AraxiVolume_VISOPRINERGY_J/Jobs/SmartHotFolders/ALL DSF JOBS HOTFOLDER/55mm x 85mm One Sided B-Cards";
                        break;

                    case "4":
                        newPath = "//VISOPRINERGY/AraxiVolume_VISOPRINERGY_J/Jobs/SmartHotFolders/ALL DSF JOBS HOTFOLDER/Black_Epsilon B-card imprints";
                        break;

                    case "5":
                        newPath = "//VISONAS/Public/DSF-Jobs/Prepress Changes 09";
                        break;

                }


                System.IO.File.Move(pathDL + "/file.pdf", pathDL + jobNum + "_BC.pdf");

                File.Copy(pathDL + jobNum + "_BC.pdf", newPath + "/" + jobNum + "_BC.pdf");

                File.Move(pathDL + jobNum + "_BC.pdf", "//VISONAS/Public/DSF-Jobs/Sent" + "/" + jobNum + "_BC.pdf");



                string description = dtTest.Rows[0]["Job-Desc"].ToString();
                string qty = dtTest.Rows[0]["Quantity-Ordered"].ToString();



                MessageBox.Show("Click Yes when finished:", "Please merge the quote within monarch", MessageBoxButtons.YesNo);

                //requery for description after merge append with description form abaove


                DataTable dtNewDesc = new DataTable();

                String sqlStatDesc = "SELECT * FROM PUB.Job WHERE \"Job-ID\" = " + jobNum;  // AND \'Sequence\' =  \'4\'";

                OdbcDataAdapter adapDesc = new OdbcDataAdapter(sqlStatDesc, dbConn);

                //fill datatable
                adapDesc.Fill(dtNewDesc);

                //append the user input ^ to the estimate description and UPDATE 
                string newDescrip = description + "\n" + dtNewDesc.Rows[0]["Job-Desc"].ToString();

                string descrip = "UPDATE PUB.Job SET \"Job-Desc\" = \'" + newDescrip + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmdMergeDesc = new OdbcCommand(descrip, dbConn);
                cmdMergeDesc.ExecuteNonQuery();
                //end set job description


                //set qty
                string qtySQL = "UPDATE PUB.Job SET \"Quantity-Ordered\" = \'" + qty + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmdMergeQTY = new OdbcCommand(qtySQL, dbConn);
                cmdMergeQTY.ExecuteNonQuery();
                //end qty



                #region Schedule board

                string transNumberPath = "//visonas/Public/Kyle/dsf job processing program/transNumber.txt";

                var number = File.ReadAllLines(transNumberPath);

                int num = Convert.ToInt32(number[0]);

                //get current date
                DateTime date = DateTime.Now;

                string time = date.ToString("HHmm");

                //set sechdule board to 3 tags and 70 
                //now for the 3 free fields 780,900,950  - this is no working for some reason? MUST USE INSERT INTO cannot update as there is nothing there

                num++;
                string SBff1 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                 " VALUES (\'" + jobNum + "\', \'780\', \'50d\', \'" + num + "\', \'1\', \'Viso\', \'730\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'780\', \'0\', \'Schedule Board\')";
                OdbcCommand sbCmd1 = new OdbcCommand(SBff1, dbConn);
                sbCmd1.ExecuteNonQuery();


                num++;
                string SBff2 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                 " VALUES (\'" + jobNum + "\', \'900\', \'50d\', \'" + num + "\', \'1\', \'Viso\', \'900\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'900\', \'0\', \'Schedule Board\')";
                OdbcCommand sbCmd2 = new OdbcCommand(SBff2, dbConn);
                sbCmd2.ExecuteNonQuery();


                num++;
                string SBff3 = "INSERT INTO PUB.ScheduleByJob (\"Job-ID\", \"Work-Center-ID\", \"TagStatus-ID\", \"Trans-Number-ScheduleByJob\", \"View-Tag\", \"System-ID\", \"SchedulingCenter-ID\", \"Date-Scheduled\", \"Update-date\", \"Created-By\", \"Prog-Name\", \"Date-Promised\", \"Time-On-In-Seconds\", \"Time-Off-In-Seconds\", \"Created-Date\", \"Update-Time\", \"Time-On\", \"Time-Off\", \"SchedulingDepartment-ID\", \"Tag-Complete\", \"Toggle1\", \"Toggle2\", \"Toggle3\", \"Number-of-Resources\",\"Date-Sort\", \"Original-Work-Center-ID\", \"Trans-Num-Task\", \"Schedule-Source\")" +
                                                 " VALUES (\'" + jobNum + "\', \'950\', \'50d\', \'" + num + "\', \'1\', \'Viso\', \'950\', \'" + date + "\', \'" + date + "\', \'kjacobsen\', \'USER-INTERFACE-TRIGGER sb/sb-sba0-d.w\', \'" + date + "\', \'35640\',\'35640\', \'" + date + "\', \'09:54:06\', \'0954\', \'0954\', \'Bin\', \'0\', \'0\',\'0\',\'0\', \'1\', \'" + date + "\', \'950\', \'0\', \'Schedule Board\')";
                OdbcCommand sbCmd3 = new OdbcCommand(SBff3, dbConn);
                sbCmd3.ExecuteNonQuery();
                #endregion Schedule Region

                //DO NOT FORGET TO write back new trans-number
                StreamWriter sw = new StreamWriter(transNumberPath);
                sw.WriteLine(num);
                sw.Close();



                Tick80DSF ticket = new Tick80DSF();

                ticket.SetDatabaseLogon("Bob", "Orchard", "monarch18", "gams1");
                ticket.SetDatabaseLogon("Bob", "Orchard");


                ticket.SetParameterValue("Job-ID", jobNum);
                ticket.SetParameterValue("System-ID", "Viso");

                crystalReportViewer1.ReportSource = ticket;
                crystalReportViewer1.Refresh();


              
            }//end connection
            }
        }//end form

}//end class

