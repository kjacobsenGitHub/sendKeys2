﻿using System;
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

                

                string transNumberPath = "C:/Users/kjacobsen/source/repos/sendKeys2/sendKeys2/transNumber.txt";

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
                #endregion Schedule Region




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
               // num++;

                StreamWriter sw = new StreamWriter(transNumberPath);
                sw.WriteLine(num);
                sw.Close();
                MessageBox.Show("Done with Inventory job "+jobNum);

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
                string expenseCode = "UPDATE PUB.Job SET \"Expense-Code\" = \'\' WHERE \"Job-ID\" = " + jobNum;
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

                string transNumberPath = "C:/Users/kjacobsen/source/repos/sendKeys2/sendKeys2/transNumber.txt";

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
                #endregion Schedule Region




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

                StreamWriter sw = new StreamWriter(transNumberPath);
                sw.WriteLine(num);
                sw.Close();
                MessageBox.Show("Done with Inventory RP job: "+jobNum);

            }//end conncetion
        }//end IRP button




        //digital job
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


                string description = dtTest.Rows[0]["Job-Desc"].ToString();

                string qty = dtTest.Rows[0]["Quantity-Ordered"].ToString();



                //now that we have decription get the quote number between ()
                string result = String.Empty;

                while (description.Contains('(') && description.Contains(')'))
                {
                    int openIndex = description.IndexOf('(');
                    int closeIndex = description.IndexOf(')');

                    result += description.Substring(openIndex + 1, closeIndex - openIndex - 1) + " ";
                    description = description.Remove(0, closeIndex + 1);
                }

                result.Trim();
         


                //need to merge quote
                //does this work? - like does it do the same as how we normally merge quotes
                string merge = "UPDATE PUB.Job SET \"Last-Estimate-ID\" = \'" + result + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmdMerge = new OdbcCommand(merge, dbConn);
                cmdMerge.ExecuteNonQuery();



                //ask user for Job-Desc first line and QTY; Job-Desc  and    Quantity-Ordered  
                string jobDesc = Interaction.InputBox("Ex) 1234 Business Card (Quote) D Kyle Jacobsen","First line Job Description:");


                string descrip = "UPDATE PUB.Job SET \"Job-Desc\" = \'" + jobDesc + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmdMergeDesc = new OdbcCommand(descrip, dbConn);
                cmdMergeDesc.ExecuteNonQuery();


                string qtySQL = "UPDATE PUB.Job SET \"Quantity-Ordered\" = \'" + qty + "\' WHERE \"Job-ID\" = " + jobNum;
                OdbcCommand cmdMergeQTY = new OdbcCommand(qtySQL, dbConn);
                cmdMergeQTY.ExecuteNonQuery();


                
                //free fields




                MessageBox.Show("Done with Digital job: " + jobNum);


            }//end connection

            }//end digital job




        //invenotry multiple
        private void button4_Click(object sender, EventArgs e)
        {


            string path = "C:/Users/kjacobsen/Desktop/jobs.txt";

            //read in txt file jobs
            var file = File.ReadAllLines(path);

            for (int x = 0; x < file.Length; x++)
            {


                string transNumberPath = "C:/Users/kjacobsen/source/repos/sendKeys2/sendKeys2/transNumber.txt";

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


           

            MessageBox.Show("Done with "+file.Length+" Invenotry jobs.");


        }//end IM


        //delete schedule transactin tags
        private void button5_Click(object sender, EventArgs e)
        {

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

    }//end form

}//end class
