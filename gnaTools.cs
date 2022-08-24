using System;
using Microsoft.Win32;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.SqlClient;
using EASendMail; //add EASendMail namespace (This needs the license code)



namespace GNAgeneraltools
{
    public class gnaTools
    {

#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8602
#pragma warning disable CS8604
#pragma warning disable CA1416




        public void HelloWorld()
        {
            Console.WriteLine("GNAgeneraltools.gnaTools : Hello world");
            Console.ReadKey();
        }

        public void WelcomeMessage(string strMessage)
        {
            Console.Title = strMessage;
            Console.WriteLine(" ");
            Console.WriteLine("GNA Geomatics software");
            Console.WriteLine("Julian Gray");
            Console.WriteLine("+49 176 7299 7904");
            Console.WriteLine("gna.geomatics@gmail.com");
            Console.WriteLine(" ");
            Console.WriteLine(strMessage);
            Console.WriteLine("");
        }

        public string checkSoftwareLicence(string strProject, string strEmailLogin, string strEmailPassword, string strSendEmail)
        {

            string strStatus = "empty";

            //
            // Check the number of days remaining on the license
            // The start date and validity period must be set using the software "SetSoftwareKey"
            //

            // This module must be at the start of all software modules
            // use:
            //
            // string strSendEmail = "No";
            // string strSoftwareKey = gna.checkSoftwareReferenceDate(strProject, strEmailLogin, strEmailPassword, strSendEmail);
            // if (strSoftwareKey == "expired") goto TheEnd;
            //
            // TheEnd:
            //  Console.WriteLine("Done");
            //
            // the answer is either "valid" or "expired"
            //  strSendEmail = "Yes" when CheckSoftwareKey is set as a scheduled task 
            //  strSendEmail is set to "No" when used inside the software
            //

            //Console.WriteLine("checkSoftwareReferenceDate");
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Diebus");
            string strValidityPeriod = key.GetValue("TempusValide", "No Value").ToString();
            key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Portunus");
            string strReferenceDate = key.GetValue("Clavis", "No Value").ToString();

            Console.WriteLine("Licence period: " + strValidityPeriod + " days");
            Console.WriteLine("Reference date: " + strReferenceDate);

            DateTime InstallDate = DateTime.ParseExact(strReferenceDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            DateTime TodayDate = DateTime.Today;

            TimeSpan interval = (TodayDate - InstallDate);
            int iRemainingDays = Convert.ToInt16(strValidityPeriod) - interval.Days;

            Console.WriteLine("Remaining days: " + iRemainingDays.ToString() + " days");

            if (iRemainingDays < 0)
            {
                strStatus = "expired";

                Console.WriteLine(" ");
                Console.WriteLine("The software license has expired.");
                Console.WriteLine("Please contact Julian Gray on +49 176 7299 7904 to reactivate");
                Console.WriteLine(" ");
                Console.ReadLine();
            }
            else
            {
                strStatus = "valid";
            }

            if ((iRemainingDays < 4) & (strSendEmail == "Yes"))
            {
                try
                {
                    SmtpMail oMail = new SmtpMail("ES-E1582190613-00131-72B1E1BD67B73FVA-C5TC1DDC612457A3");
                    {
                        oMail.From = "gna.geomatics@gmail.com";
                        oMail.To = new AddressCollection("gna.geomatics@gmail.com");
                    };

                    // Set email subject
                    oMail.Subject = "Software license about to expire (" + strProject + ")";
                    // Set email body
                    oMail.TextBody = "No of days remaining: " + iRemainingDays.ToString();

                    // SMTP server address
                    SmtpServer oServer = new SmtpServer("smtp.gmail.com")
                    {
                        User = strEmailLogin,
                        Password = strEmailPassword,
                        ConnectType = SmtpConnectType.ConnectTryTLS,
                        Port = 587
                    };

                    SmtpClient oSmtp = new SmtpClient();
                    oSmtp.SendMail(oServer, oMail);

                    Console.WriteLine("Advisory email issued");

                }
                catch (Exception ep)
                {
                    Console.WriteLine("Email transmission failed..");
                    Console.WriteLine(strEmailLogin);
                    Console.WriteLine(strEmailPassword);
                    Console.WriteLine("");
                    Console.WriteLine(ep.Message);
                    Console.ReadKey();
                }
            }

            return strStatus;
        }


        public string checkLicenseValidity(string strSoftwareLicenseTag, string strProject, string strEmailLogin, string strEmailPassword, string strSendEmail)
        {


            string strStatus = "empty";

            //
            // Check the number of days remaining on the license
            // The start date and validity period must be set using the software "LicenseCreator"
            // The software license tag used in creating the license must exist in the parent software and be passed through to this method as a parameter

            // This module must be at the start of all software modules
            // use:
            //
            // string strSendEmail = "No";
            // string strSoftwareKey = gna.checkSoftwareReferenceDate(strProject, strEmailLogin, strEmailPassword, strSendEmail);
            // if (strSoftwareKey == "expired") goto TheEnd;
            //
            // TheEnd:
            //  Console.WriteLine("Done");
            //
            // the answer is either "valid" or "expired"
            //  strSendEmail = "Yes" when CheckSoftwareKey is set as a scheduled task 
            //  strSendEmail is set to "No" when used inside the software
            //

            string strDurationSubKey = @"SOFTWARE\Diebus_" + strSoftwareLicenseTag;
            RegistryKey key = Registry.CurrentUser.OpenSubKey(strDurationSubKey);
            string strValidityPeriod = key.GetValue("TempusValide", "No Value").ToString();

            string strReferenceDateSubKey = @"SOFTWARE\Portunus_" + strSoftwareLicenseTag;
            key = Registry.CurrentUser.OpenSubKey(strReferenceDateSubKey);
            string strReferenceDate = key.GetValue("Clavis", "No Value").ToString();

            Console.WriteLine("License start date: " + strReferenceDate);
            Console.WriteLine("Licence validity: " + strValidityPeriod + " days");

            DateTime InstallDate = DateTime.ParseExact(strReferenceDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            DateTime TodayDate = DateTime.Today;

            TimeSpan interval = (TodayDate - InstallDate);
            int iRemainingDays = Convert.ToInt16(strValidityPeriod) - interval.Days;

            Console.WriteLine("Remaining: " + iRemainingDays.ToString() + " days");

            if (iRemainingDays < 0)
            {
                strStatus = "expired";

                Console.WriteLine(" ");
                Console.WriteLine("The software license has expired.");
                Console.WriteLine("Please contact Julian Gray to reactivate:");
                Console.WriteLine("gna.geomatics@gmail.com");
                Console.WriteLine("+49 176 7299 7904");
                Console.WriteLine(" ");
                Console.ReadLine();
            }
            else
            {
                strStatus = "valid";
            }

            if ((iRemainingDays < 4) & (strSendEmail == "Yes"))
            {
                try
                {
                    SmtpMail oMail = new SmtpMail("ES-E1582190613-00131-72B1E1BD67B73FVA-C5TC1DDC612457A3");
                    {
                        oMail.From = "gna.geomatics@gmail.com";
                        oMail.To = new AddressCollection("gna.geomatics@gmail.com");
                    };

                    // Set email subject
                    oMail.Subject = "Software license (" + strProject + ")";
                    // Set email body
                    oMail.TextBody = "No of days remaining: " + iRemainingDays.ToString();

                    // SMTP server address
                    SmtpServer oServer = new SmtpServer("smtp.gmail.com")
                    {
                        User = strEmailLogin,
                        Password = strEmailPassword,
                        ConnectType = SmtpConnectType.ConnectTryTLS,
                        Port = 587
                    };

                    SmtpClient oSmtp = new SmtpClient();
                    oSmtp.SendMail(oServer, oMail);

                    Console.WriteLine("Advisory email issued");

                }
                catch (Exception ep)
                {
                    Console.WriteLine("Email transmission failed..");
                    Console.WriteLine(strEmailLogin);
                    Console.WriteLine(strEmailPassword);
                    Console.WriteLine("");
                    Console.WriteLine(ep.Message);
                    Console.ReadKey();
                }
            }

            return strStatus;
        }


        //=================[ Date Time methods ]======================================================================================================

        public string convertLocalToUTC(string strLocalTime)
        {
            // Purpose: 
            //      To convert the local time string to UTC time string
            // Input: 
            //      strLocalTime: '2022-08-01 00:00'
            // Output:
            //      strUTCtime: '2022-08-01 00:00'
            // Use:
            //      strUTCtime = convertLocalToUTC(strLocalTime);
            // 

            // strip off the ' if they exist
            strLocalTime = strLocalTime.Replace("'","").Trim();

            DateTime dtNowLocal = DateTime.Now;
            DateTime dtNowUTC = DateTime.UtcNow;

            DateTime dtLocalTime = DateTime.ParseExact(strLocalTime, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

            double dblHoursOffset = Math.Round(((dtNowUTC-dtNowLocal).Minutes)/60.0,0);

            DateTime dtUTCtime = dtLocalTime.AddHours(dblHoursOffset);

            string strUTCtime = " '" + dtUTCtime.ToString("yyyy-MM-dd HH:mm")+"'";

            return strUTCtime;
        }

        public string[,] generateTimeBlockArray(int iTimeBlockSize, int iNoOfDays)
        { 
            //
            // Purpose:
            //      To create an array of start/end time blocks between Now and back iNoOfDays
            // Input:
            //      No of days back from today (Rolling window)
            //      Time block size (hours)
            //      No of time blocks
            // Output:
            //      Returns strTimeBlockArray [counter, start time, end time] [iCounter,0-1]
            // Useage:
            //      int iBlockSize = Convert.ToInt32(strBlockSize);
            //      int iNoOfDays = Convert.ToInt16(strRollingTimeWindowDays); ;
            //      string[,] strTimeBlockArray = gna.generateTimeBlockArray(iBlockSize, iNoOfDays);
            //      string[iCounter,0] : strCounter
            //      string[iCounter,1] : strStartTime
            //      string[iCounter,2] : strEndTime
            // Comment:
            //      Last time = "NoMore"
            //

            //double dblTimeOffset = Convert.ToDouble(iOffsetFromNow * (-1.0));
            double dblTimeBlockSize = Convert.ToDouble(iTimeBlockSize);
            //DateTime dtUTCnow = DateTime.UtcNow;
            //DateTime dtEndTime = DateTime.UtcNow;
            //DateTime dtStartTime = DateTime.UtcNow;

            string[,] strTimeBlockArray = new string[2000, 3];

            // Compute the start date/time (local time)
            DateTime dtToday = DateTime.Today;
            double dblDaysOffset = Convert.ToDouble(iNoOfDays * (-1.0));
            double dblHoursOffset = 100;
            DateTime dtStartDate = dtToday.AddDays(dblDaysOffset);
            DateTime dtNow = DateTime.Now;
            DateTime dtStartTime = dtStartDate;
            DateTime dtEndTime = dtStartTime.AddHours(dblTimeBlockSize);

            int iCounter = 0;

            do
            {
                dblHoursOffset = (dtNow - dtEndTime).TotalHours;
                string strStartTime = "'" + dtStartTime.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                string strEndTime = "'" + dtEndTime.ToString("yyyy-MM-dd HH:mm:ss") + "'";

                strTimeBlockArray[iCounter, 0] = Convert.ToString(iCounter);
                strTimeBlockArray[iCounter, 1] = strStartTime;
                strTimeBlockArray[iCounter, 2] = strEndTime;
                iCounter++;

                dtStartTime = dtEndTime;
                dtEndTime = dtStartTime.AddHours(dblTimeBlockSize);
                dblHoursOffset = (dtNow - dtEndTime).TotalHours;

            } while (dblHoursOffset >= 0);

            strTimeBlockArray[iCounter, 0] = "NoMore";
            strTimeBlockArray[iCounter, 1] = "2000-01-01 00:00:00";
            strTimeBlockArray[iCounter, 2] = "2000-01-01 01:00:00";
            return strTimeBlockArray;
        }


        public Tuple<string, string> generateTimeBlockStartEnd(int iOffsetFromNow, int iTimeBlockSize, int iTimeBlockNumber)
        {
            // Purpose: 
            //      To compute the start and end times of the time block used to extract data from the DB
            // Input: 
            //      OffsetFromNow, TimeBlockSize, TimeBlockNumber (hours,integer,integer)
            // Output:
            //      BlockStartTime, BlockEndTime
            // Use:
            //      App.config gives the initial time offset, block size and number of blocks
            //      For time block number from 0 to number of blocks
            //          extract the start and end times
            //              var answer = gna.TimeBlockStartEnd(2, 4, 0);
            //              string strTimeBlockStartUTC = answer.Item1;
            //              string strTimeBlockEndUTC = answer.Item2;
            //          extract the data from the DB within those start and end times
            //          update the worksheet

            iOffsetFromNow = iOffsetFromNow + ((iTimeBlockNumber - 1) * iTimeBlockSize);
            double dblStartTimeOffset = -1.0 * (Convert.ToDouble(iOffsetFromNow));
            double dblEndTimeOffset = dblStartTimeOffset - (Convert.ToDouble(iTimeBlockSize));

            string strTimeBlockStartUTC = " '" + DateTime.UtcNow.AddHours(dblEndTimeOffset).ToString("yyyy-MM-dd HH:mm") + "' ";
            string strTimeBlockEndUTC = " '" + DateTime.UtcNow.AddHours(dblStartTimeOffset).ToString("yyyy-MM-dd HH:mm") + "' ";
            return new Tuple<string, string>(strTimeBlockStartUTC, strTimeBlockEndUTC);
        }


        public double UTCtimeOffset()
        {
            //
            // Purpose:
            //      To generate the UTC time offset for the server
            // Input:
            //      Nothing
            // Output:
            //      The difference between UTC and now in double hours
            // Useage:
            //      double dblUTCoffset = gna.UTCtimeOffset();
            //

            double dblUTCtimeOffset = 0;

            DateTime dtUTCnow = DateTime.UtcNow;
            DateTime dtNow = DateTime.Now;
            dblUTCtimeOffset = Convert.ToDouble((dtUTCnow - dtNow).Hours);

            return dblUTCtimeOffset;
        }

        public string formatTimestampMissionOS(string strTimestamp)
        {
            // 
            // Formats the time stamp as required by MissionOS
            // Receives 19/03/2021 11:43:17 (UK setting)
            // Returns 2021-03-19 11:43:17 (MissionOS)
            // 
            // strTimeStamp = gna.formatTimestampMissionOS(strTimeStamp);
            //


            strTimestamp = strTimestamp.Trim();

            // Now reformat this 
            string YYYY = strTimestamp.Substring(6, 4);
            string MM = strTimestamp.Substring(3, 2);
            string dd = strTimestamp.Substring(0, 2);
            string time = strTimestamp.Substring(11, 8);
            strTimestamp = YYYY + "-" + MM + "-" + dd + " " + time;
            return strTimestamp;
        }


        public Tuple<string, string> generateHistoricTimeBlockStartEnd(string strHistoricTimeBlockStart, string strHistoricTimeBlockEnd, int iHistoricTimeBlockSize, int iTimeBlockCounter)
        {
            // Purpose: 
            //      To compute the start and end times of the time block used to by the historic data function in data extraction
            // Input: 
            //      strHistoricTimeBlockStart, strHistoricTimeBlockEnd, iHistoricTimeBlockSize, iTimeBlockCounter
            // Output:
            //      strTimeBlockStartUTC, strTimeBlockEndUTC
            //      if the time block goes beyond the end date, it returns "The End"
            // Use:
            //      var answer = gnaT.generateHistoricTimeBlockStartEnd(strHistoricTimeBlockStart, strHistoricTimeBlockEnd, iHistoricTimeBlockSize,iTimeBlockCounter);
            //      string strTimeBlockStartUTC = answer.Item1;
            //      string strTimeBlockEndUTC = answer.Item2;
            //      extract the data from the DB within those start and end times
            //         

            string strTimeBlockStartLocal = "";
            string strTimeBlockEndLocal = "";
            string strTimeBlockStartUTC = "";
            string strTimeBlockEndUTC = "";

            // strip off the ' if they exist
            strHistoricTimeBlockStart = strHistoricTimeBlockStart.Replace("'", "").Trim();
            strHistoricTimeBlockEnd = strHistoricTimeBlockEnd.Replace("'", "").Trim();

            DateTime dtHistoricTimeStart = DateTime.ParseExact(strHistoricTimeBlockStart, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
            DateTime dtHistoricTimeEnd = DateTime.ParseExact(strHistoricTimeBlockEnd, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

            double dblOffset1 = Convert.ToDouble(iHistoricTimeBlockSize * (iTimeBlockCounter-1));
            double dblOffset2 = Convert.ToDouble(iHistoricTimeBlockSize * (iTimeBlockCounter));

            DateTime dtCurrentTimeStart = dtHistoricTimeStart.AddHours(dblOffset1);
            DateTime dtCurrentTimeEnd = dtHistoricTimeStart.AddHours(dblOffset2);

            if (DateTime.Compare(dtHistoricTimeEnd, dtCurrentTimeStart) == 0)
            {
                strTimeBlockStartUTC = "The End";
                strTimeBlockEndUTC = "The End";
            }
            else
            {
                strTimeBlockStartLocal = " '" + dtCurrentTimeStart.ToString("yyyy-MM-dd HH:mm") + "' ";
                strTimeBlockEndLocal = " '" + dtCurrentTimeEnd.ToString("yyyy-MM-dd HH:mm") + "' ";
                strTimeBlockStartUTC = convertLocalToUTC(strTimeBlockStartLocal);
                strTimeBlockEndUTC = convertLocalToUTC(strTimeBlockEndLocal);
            }
            return new Tuple<string, string>(strTimeBlockStartUTC, strTimeBlockEndUTC);
        }


    }
}