using System.Globalization;
using System.Timers;

using gnaDataClasses;

using EASendMail; //add EASendMail namespace (This needs the license code till 2028: ES-E1646458156-01989-1A16E5275AF48A22-686E917424789B4F)
using Microsoft.Win32;
using Twilio.Rest.Api.V2010.Account;




#pragma warning disable CS8600
#pragma warning disable IDE0059
#pragma warning disable IDE1006
#pragma warning disable CA1822
#pragma warning disable CS8602
#pragma warning disable CS8603
#pragma warning disable CS8604
#pragma warning disable CS8622
#pragma warning disable CA1416



namespace GNAgeneraltools
{

    public class gnaTools
    {
        //================[actions]=====================
        // 20231010 Create CountLinesLINQ()
        // 20231010 Create numberOfLinesInTextFile()
        // 20231025 Added options to the sms phones
        // 20240206 Updated copyright year
        // 20240429 Create updateSystemAlarmFile
        // 20240509 Create updateNoDataAlarmFile
        // 20240522 Creaate doesFileExist
        // 20240531 Added the software name to the copyright notice


        //=====[Notes]=============
        // easier way to sort a list
        //
        // instantiate the list
        // var car = new List<Car>();
        // car = car.OrderBy(x => x.Model).ToList();

        //================[instantiate the classes]=========
        gnaDataClass gnaDC = new();


        //================[File management Methods ]========

        public Int32 CountLinesLINQ(FileInfo file) => File.ReadLines(file.FullName).Count();

        public Int32 numberOfLinesInTextFile(FileInfo file)
        {

            Int32 iNoOfLines = CountLinesLINQ(file);

            return iNoOfLines;
        }

        //================[ Coordinate Methods ]========

        public void checkCoordinates(List<Prism> prism)
        {

            prism.Sort(delegate (Prism x, Prism y)
            {
                return x.Name.CompareTo(y.Name);
            });

            int iPrismCount = prism.Count;
            string strDuplicateFound = "No";

            for (int i = 0; i < iPrismCount - 1; i++)
            {
                if (prism[i].Name.Trim() == prism[i + 1].Name.Trim())
                {
                    Console.WriteLine("      Duplicate name found: " + prism[i].Name);
                    strDuplicateFound = "Yes";
                }
            }

            prism.Sort((Prism x, Prism y) =>
            {
                int result = x.E.CompareTo(y.E);
                return result == 0 ? x.N.CompareTo(y.N) : result;
            });

            for (int i = 0; i < iPrismCount - 1; i++)
            {

                Double dblDist = Math.Round(Math.Pow(Math.Pow(prism[i + 1].E - prism[i].E, 2) + Math.Pow(prism[i + 1].N - prism[i].N, 2), 0.5), 2);

                if (dblDist < 0.05)
                {
                    Console.WriteLine("      Duplicate coordinates found: " + prism[i].Name + " : " + prism[i + 1].Name);
                    strDuplicateFound = "Yes";
                }
            }

            if (strDuplicateFound == "Yes")
            {
                Console.WriteLine("\n   Fix duplicates before continuing");
                Console.WriteLine("   Press key to stop..");
                Console.ReadKey();
                Environment.Exit(0);
            }

        }


        //================[ SMS Methods ]===============

        public void checkReportingSchedule(string strTimeBlockType, string strSMSTitle)
        {

            if (strTimeBlockType == "Manual")
            {
                // send a warning sms
                string strMessage = strSMSTitle + "\nManual time schedule";
                sendSMS(strMessage, "None", "None", "Yes");
                Console.WriteLine("     SMS warning issued");
            }
            else
            {

                Console.WriteLine("     No warning issued");
            }
        }

        //================[ General Methods ]===========

        public static void HelloWorld()
        {
            Console.WriteLine("GNAgeneraltools.gnaTools : Hello world");
        }

        private static void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            Console.WriteLine("System timeout...");
            Environment.Exit(0);
        }


        public void freezeScreen(string strFreezeScreen)
        {

            // Create a timer with a 20 second interval.
            System.Timers.Timer aTimer = new();
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.Interval = 20000;
            aTimer.Enabled = true;

            if (strFreezeScreen == "Yes")
            {
                Console.WriteLine("\nfreezeScreen: Yes \nPress key to exit..");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("\nNo freezeScreen");
            }
        }

        public string check2wayComms(string strIncomingMessage)
        {
            string strOutgoingMessage = strIncomingMessage + " (Received)";
            return strOutgoingMessage;
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


        public string addCopyright(string software, string strIncomingMessage)
        {
            string strCopyRight = "\n©2024 GNA Geomatics";
            software = software.Trim();
            if (software.Length > 0)
            {
                strCopyRight = "\n" + software + strCopyRight;
            }
            string strOutgoingMessage = strIncomingMessage + strCopyRight;
            return strOutgoingMessage;
        }

        //=================[ Registry methods ]====================================================


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

            TimeSpan interval = TodayDate - InstallDate;
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
                    SmtpMail oMail = new("ES-E1646458156-01989-1A16E5275AF48A22-686E917424789B4F")
                    {
                        From = "gna.geomatics@gmail.com",
                        To = new AddressCollection("gna.geomatics@gmail.com"),
                        Subject = "Software license about to expire (" + strProject + ")",
                        TextBody = "No of days remaining: " + iRemainingDays.ToString()
                    };



                    // SMTP server address
                    SmtpServer oServer = new("smtp.gmail.com")
                    {
                        User = strEmailLogin,
                        Password = strEmailPassword,
                        ConnectType = SmtpConnectType.ConnectTryTLS,
                        Port = 587
                    };

                    SmtpClient oSmtp = new();
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

            string strDurationSubKey = "";
            string strSoftwareValidityPeriod = "";
            string strReferenceDateSubKey = "";
            string strReferenceDate = "";
            string software = "";

            switch (strSoftwareLicenseTag)
            {
                case "CIV177":
                    software = "civ177";
                    break;
                case "BBVTGR":
                    software = "BBV";
                    break;

                case "EXPCRD":
                    software = "exptCrds";
                    break;
                case "SVRWDG":
                    software = "svrWatchdog";
                    break;
                case "CIV177TGR":
                    software = "civ177";
                    break;
                case "DSPRPT":
                    software = "dsplRpt";
                    break;
                case "SPN010":
                    software = "SPN010";
                    break;
                case "PJTPFM":
                    software = "pjtPfm";
                    break;
                case "FTP":
                    software = "FTP";
                    break;
                case "GKADAT":
                    software = "GKAtoDAT";
                    break;
                case "WDGALM":
                    software = "watchdogAlm";
                    break;
                case "DATALM":
                    software = "dataFlwAlm";
                    break;
                case "SETTOP":
                    software = "settopOutOfToleranceAlarm";
                    break;
                default:
                    software = "";
                    break;
            }



            if (strSoftwareLicenseTag == "")
            {
                Console.WriteLine("\n There is no license for this software");
                Console.WriteLine("Press key..."); Console.ReadKey();
                Console.ReadKey();
                Environment.Exit(0);
            }

            try
            {

                // To read the software reference date
                strReferenceDateSubKey = @"SOFTWARE\Portunus_" + strSoftwareLicenseTag;
                strDurationSubKey = @"SOFTWARE\Diebus_" + strSoftwareLicenseTag;

                RegistryKey key = Registry.CurrentUser.OpenSubKey(strDurationSubKey);
                strSoftwareValidityPeriod = key.GetValue("TempusValide", "No Value").ToString();

                key = Registry.CurrentUser.OpenSubKey(strReferenceDateSubKey);
                strReferenceDate = key.GetValue("Clavis", "No Value").ToString();
            }
            catch
            {
                Console.WriteLine("\nThere is no license for this software");
                Console.WriteLine("\nPress key..."); Console.ReadKey();
                Console.ReadKey();
                Environment.Exit(0);
            }




            DateTime InstallDate = DateTime.ParseExact(strReferenceDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            DateTime TodayDate = DateTime.Today;

            TimeSpan interval = TodayDate - InstallDate;
            int iRemainingDays = Convert.ToInt16(strSoftwareValidityPeriod) - interval.Days;

            if (iRemainingDays < 0)
            {
                strStatus = "expired";

                strReferenceDateSubKey = @"SOFTWARE\Portunus_" + strSoftwareLicenseTag;
                strDurationSubKey = @"SOFTWARE\Diebus_" + strSoftwareLicenseTag;


                Registry.CurrentUser.DeleteSubKey(strReferenceDateSubKey, false);
                Registry.CurrentUser.DeleteSubKey(strDurationSubKey, false);

                Console.WriteLine("\nThe software license for "+ software+"  has expired.");
                Console.WriteLine("Contact Julian Gray to reactivate:");
                Console.WriteLine("gna.geomatics@gmail.com");
                Console.WriteLine("+49 176 7299 7904");
                Console.WriteLine(" ");
                Console.ReadLine();
            }
            else
            {
                strStatus = "valid";
                Console.WriteLine("License validated");
                Console.WriteLine("Remaining period: " + iRemainingDays + " days");
            }

            if ((iRemainingDays < 4) & (strSendEmail == "Yes"))
            {

                try
                {
                    string strMessage = strProject + " (" + strSoftwareLicenseTag + ")\nRemaining days: " + iRemainingDays.ToString();
                    strMessage = addCopyright(software,strMessage);

                    SmtpMail oMail = new("ES-E1646458156-01989-1A16E5275AF48A22-686E917424789B4F")
                    {

                        From = "gna.geomatics@gmail.com",
                        To = new AddressCollection("gna.geomatics@gmail.com"),
                        // Set email subject
                        Subject = "Software license (" + strProject + ")",
                        // Set email body
                        TextBody = strMessage
                    };
                    // SMTP server address
                    SmtpServer oServer = new("smtp.gmail.com")
                    {
                        User = strEmailLogin,
                        Password = strEmailPassword,
                        ConnectType = SmtpConnectType.ConnectTryTLS,
                        Port = 587
                    };

                    SmtpClient oSmtp = new();
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


        public void updateReportTime(string strReportIdentifier)
        {
            // To set the time of the latest report in the registry
            //Console.WriteLine("Report Tag: " + strReportIdentifier);

            string strSoftwareTagSubKey = @"SOFTWARE\Report_" + strReportIdentifier;
            Registry.CurrentUser.DeleteSubKey(strSoftwareTagSubKey, false);
            RegistryKey rk = Registry.CurrentUser.CreateSubKey(strSoftwareTagSubKey);

            // Create name/value pairs.Time of report is stored in variable ReportTime
            string strNow = DateTime.UtcNow.ToString("yyyy-MM-dd HH");

            rk.SetValue("ReportTime", strNow);

            //Console.WriteLine("software tag: "+strSoftwareTagSubKey);
            //Console.WriteLine("ReportTime: "+ strNow);

        }

        public string checkReportTime(string strReportIdentifier, int iTimeWindowHrs)
        {

            string strReportStatus = strReportIdentifier + "Failed";

            try
            {
                // To read the latest successful report time
                string strSoftwareTagSubKey = @"SOFTWARE\Report_" + strReportIdentifier;
                RegistryKey key = Registry.CurrentUser.OpenSubKey(strSoftwareTagSubKey);
                string strLastSuccessfulReport = key.GetValue("ReportTime", "No Value").ToString();
                string strNow = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss");

                DateTime dtLastSuccessfulReport = DateTime.Parse(strLastSuccessfulReport + ":00:00", CultureInfo.InvariantCulture);
                DateTime dtNow = DateTime.Parse(strNow, CultureInfo.InvariantCulture);

                double dblHours = Convert.ToDouble((dtNow - dtLastSuccessfulReport).Days) * 24.0;
                dblHours += Convert.ToDouble((dtNow - dtLastSuccessfulReport).Hours);


                //Console.WriteLine(dtLastSuccessfulReport);
                //Console.WriteLine(dtNow);
                //Console.WriteLine("Hrs "+dblHours);
                //Console.WriteLine(iTimeWindowHrs);

                strReportStatus = dblHours > iTimeWindowHrs ? strLastSuccessfulReport : "OK";
            }
            catch
            {
                Console.WriteLine("Error: checkReportTime: " + strReportIdentifier);
                Console.WriteLine("\nThere is no registry value for this report");
                Console.WriteLine("\nPress key..."); Console.ReadKey();
                Console.ReadKey();
                Environment.Exit(0);
            }

            return strReportStatus;
        }



        //=================[ Date Time methods ]==============================================

        public int daysBetweenDates(string strStartDate, string strEndDate)
        {
            // Purpose: 
            //      To compute the number of days between 2 dates
            // Input: 
            //      strStartDate: '2022-08-01 00:00'
            //      strEndDate: 2022-08-01 00:00:00  - any option
            // Output:
            //      iNoOfDays : integer
            // Use:
            //      iNoOfDays = daysBetweenDates(strStartDate, strEndDate);
            // 

            // strip off the ' if they exist


            //Console.WriteLine("local time: " + strLocalTime);
            //Console.WriteLine("Press key..."); Console.ReadKey();

            strStartDate = strStartDate.Replace("'", "").Trim();
            strEndDate = strEndDate.Replace("'", "").Trim();

            //DateTime dtNowLocal = DateTime.Now;
            //DateTime dtNowUTC = DateTime.UtcNow;

            //DateTime dtLocalTime = DateTime.ParseExact(strLocalTime, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

            DateTime dtStartDate = DateTime.Parse(strStartDate, CultureInfo.InvariantCulture);
            DateTime dtEndDate = DateTime.Parse(strEndDate, CultureInfo.InvariantCulture);

            Int16 iDays = Convert.ToInt16((dtEndDate - dtStartDate).Days);


            return iDays + 1;
        }


        public string convertUTCToLocal(string strUTCTime)
        {
            // Purpose: 
            //      To convert the UTC time string to Local time string
            // Input: 
            //      strUTCTime: '2022-08-01 00:00'
            // Output:
            //      strLocaltime: '2022-08-01 00:00'
            // Use:
            //      strLocaltime = convertUTCToLocal(strUTCTime);
            // 

            // strip off the ' if they exist


            //Console.WriteLine("local time: " + strLocalTime);
            //Console.WriteLine("Press key..."); Console.ReadKey();

            strUTCTime = strUTCTime.Replace("'", "").Trim();

            DateTime dtNowLocal = DateTime.Now;
            DateTime dtNowUTC = DateTime.UtcNow;

            //DateTime dtUTCTime = DateTime.ParseExact(strUTCTime, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

            DateTime dtUTCTime = DateTime.Parse(strUTCTime, CultureInfo.InvariantCulture);

            double dblHoursOffset = Math.Round((dtNowLocal - dtNowUTC).Minutes / 60.0, 0);

            DateTime dtLocaltime = dtUTCTime.AddHours(dblHoursOffset);

            string strLocaltime = " '" + dtLocaltime.ToString("yyyy-MM-dd HH:mm") + "'";

            return strLocaltime;
        }





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


            //Console.WriteLine("local time: " + strLocalTime);
            //Console.WriteLine("Press key..."); Console.ReadKey();

            strLocalTime = strLocalTime.Replace("'", "").Trim();

            DateTime dtNowLocal = DateTime.Now;
            DateTime dtNowUTC = DateTime.UtcNow;

            //DateTime dtLocalTime = DateTime.ParseExact(strLocalTime, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

            DateTime dtLocalTime = DateTime.Parse(strLocalTime, CultureInfo.InvariantCulture);

            double dblHoursOffset = Math.Round((dtNowUTC - dtNowLocal).Minutes / 60.0, 0);

            DateTime dtUTCtime = dtLocalTime.AddHours(dblHoursOffset);

            string strUTCtime = " '" + dtUTCtime.ToString("yyyy-MM-dd HH:mm") + "'";

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

            iOffsetFromNow += (iTimeBlockNumber - 1) * iTimeBlockSize;
            double dblStartTimeOffset = -1.0 * Convert.ToDouble(iOffsetFromNow);
            double dblEndTimeOffset = dblStartTimeOffset - Convert.ToDouble(iTimeBlockSize);

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


            strTimestamp = strTimestamp.Trim().Replace("'", "");

            string[] strList1 = new string[4];

            String[] strlist = strTimestamp.Split(" ", 2);  // split: [0]=19/03/2021 [1]=11:43:17
            int i = 0;
            foreach (String s in strlist)
            {
                strList1[i] = s.Trim();
                i++;
            }
            string time = strList1[1];
            strlist = strList1[0].Split("/", 3);  // split: [0]=19/03/2021

            i = 0;
            foreach (String s in strlist)
            {
                strList1[i] = s.Trim();
                i++;
            }
            string YYYY = strList1[2];
            string MM = strList1[1];
            string dd = strList1[0];

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

            double dblOffset1 = Convert.ToDouble(iHistoricTimeBlockSize * (iTimeBlockCounter - 1));
            double dblOffset2 = Convert.ToDouble(iHistoricTimeBlockSize * iTimeBlockCounter);

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


        //================[ Alarm Methods ]=================================================================

        // create an alarm log file to upate the Alarm Log File - seperate updating the file and sending emails.



        public List<ATS> updateAlarmStatus(string strSystemLogFolder, List<ATS> ats)
        {
            // Purpose
            //  To update the AlarmStatusFile and trigger an advisory email if needed
            // Input
            //     strSystemStatusLogsFolder where the AlarmStatusFile.txt is stored
            //     list of ATS
            // Return
            //      Current state with (Yes) or (No) if the state has changed

            var currentState = new List<string>();
            int j = 0;


            string strSendEmail = "";

            string strFileName = "AlarmStatus.txt";
            string alarmFile = strSystemLogFolder + strFileName;
            checkFileExistance(strSystemLogFolder, strFileName);


            // Read the current state of the system
            currentState = File.ReadAllLines(alarmFile).ToList();
            currentState.Add("TheEnd");

            // write the new state to the text file
            File.Delete(alarmFile);
            using (StreamWriter writetext = new(alarmFile, false))
            {
                j = 0;
                do
                {
                    writetext.WriteLine(ats[j].Name + "(" + ats[j].Settop + ") " + ats[j].Note1);
                    j++;
                } while (ats[j].Name != "TheEnd");
                writetext.Close();
            }

            //j = 0;
            //Console.WriteLine("\nCurrent state");
            //do
            //{
            //    Console.WriteLine(currentState[j]);
            //    j++;

            //} while (currentState[j] != "TheEnd");

            j = 0;
            do
            {
                string subString = ats[j].Name;
                int k = 0;
                string strStateChange = "";
                do
                {
                    if (currentState[k].Contains(subString))
                    {
                        if (currentState[k].Contains("OK"))
                        {
                            strStateChange = ats[j].Note1 == "Alarm" ? "Alarm (Yes)" : "OK (No)";
                        }
                        else if (currentState[k].Contains("Alarm"))
                        {
                            strStateChange = ats[j].Note1 == "OK" ? "OK (Yes)" : "Alarm (No)";
                        }
                    }
                    k++;
                } while (currentState[k] != "TheEnd");

                ats[j].Note1 = strStateChange;

                j++;

            } while (ats[j].Name != "TheEnd");


            //j = 0;
            //Console.WriteLine("\nNew state");
            //do
            //{
            //    Console.WriteLine(ats[j].Name + " : " + ats[j].Note1);
            //    j++;

            //} while (ats[j].Name != "TheEnd");

            return ats;
        }

        public void emailSystemStatus(string strProjectTitle, string[] strSettop, string[] strGKAfolder, string strEmailLogin, string strEmailPassword, string strEmailFrom, string strSystemStatusRecipients)
        {
            //
            // This is a daily email that reports, in 1 email, the state of every Settop
            // It also verifies that the alarm function is working
            // The log file for each Settop is updated to confirm that the system status email was sent
            //

            string strFolder = "";
            string strSettopNumber = "";
            string strAlarmState = "";
            string[] strStatus = new string[7];
            int iSettopCounter = 0;
            string strSubjectLine = strProjectTitle + ": System OK";

            for (int i = 1; i < 11; i++)
            {

                strFolder = strGKAfolder[i];
                strSettopNumber = strSettop[i];

                if (strFolder == "None")
                {
                    break;
                }

                string alarmFile = strFolder + "AlarmFile.txt";
                strAlarmState = File.ReadAllText(alarmFile).TrimEnd();
                if (strAlarmState == "0")
                {
                    strAlarmState = "OK";
                    strSubjectLine = strProjectTitle + ": System OK";
                }
                else
                {
                    strAlarmState = "Alarm";
                    strSubjectLine = strProjectTitle + ": System in Alarm state";
                }
                strStatus[i] = strSettopNumber + ": " + strAlarmState;
                iSettopCounter = i;
                updateAlarmLogFile(strGKAfolder[i], "System status email sent");
            }

            string strEmailBody = "\r\n";

            for (int i = 1; i <= iSettopCounter; i++)
            {
                strEmailBody = strEmailBody + strStatus[i] + "\r\n";
            }
            sendEmail(strEmailLogin, strEmailPassword, strEmailFrom, strSystemStatusRecipients, strSubjectLine, strEmailBody);
            Console.WriteLine("Status email sent");
            //Console.ReadKey();
        }

        public string updateAlarmStatusFile(string strSystemLogFolder, string strSystemAlarmState)
        {

            // NO LONGER USED


            // Purpose
            //  To update the AlarmStatusFile and trigger an advisory email if needed
            // Input
            //     strSystemStatusLogsFolder where the AlarmStatusFile.txt is stored
            //     strAlarmState Alarm/OK
            // Return
            //     sendEmail = "No"
            //     sendEmail = "Alarm"
            //     sendEmail = "SystemResetToOK"
            //

            string strSendEmail = "";

            string strFileName = "AlarmStatusFile.txt";
            string alarmFile = strSystemLogFolder + strFileName;
            checkFileExistance(strSystemLogFolder, strFileName);


            //Retrieve the current state (0,1,2,3)
            string strAlarmStateInFile = File.ReadAllText(alarmFile).TrimEnd();  // retrieving 0,1,2,3 from the Alarm log file AlarmFile.txt

            switch (strAlarmStateInFile)
            {
                case "0":
                    if (strSystemAlarmState == "OK")
                    {
                        strSendEmail = "No";
                        strAlarmStateInFile = "0";
                    }
                    else
                    {
                        strSendEmail = "Alarm1";
                        strAlarmStateInFile = "1";
                    }
                    break;
                case "1":
                    if (strSystemAlarmState == "OK")
                    {
                        strSendEmail = "SystemResetToOK";
                        strAlarmStateInFile = "0";
                    }
                    else
                    {
                        strSendEmail = "Alarm2";
                        strAlarmStateInFile = "2";
                    }
                    break;
                case "2":
                    if (strSystemAlarmState == "OK")
                    {
                        strSendEmail = "SystemResetToOK";
                        strAlarmStateInFile = "0";
                    }
                    else
                    {
                        strSendEmail = "Alarm3";
                        strAlarmStateInFile = "3";
                    }
                    break;
                case "3":
                    if (strSystemAlarmState == "OK")
                    {
                        strSendEmail = "SystemResetToOK";
                        strAlarmStateInFile = "0";
                    }
                    else
                    {
                        strSendEmail = "No";
                        strAlarmStateInFile = "4";
                    }
                    break;
                case "4":
                    if (strSystemAlarmState == "OK")
                    {
                        strSendEmail = "SystemResetToOK";
                        strAlarmStateInFile = "0";
                    }
                    else
                    {
                        strSendEmail = "No";
                        strAlarmStateInFile = "4";
                    }
                    break;
                default:
                    strSendEmail = "No";
                    strAlarmStateInFile = "0";
                    break;
            }

            File.Delete(alarmFile);
            using (StreamWriter writetext = new(alarmFile, false))
            {
                writetext.WriteLine(strAlarmStateInFile);
                writetext.Close();
            }

            return strSendEmail;
        }

        public void updateAlarmLogFile(string strCurrentGKAfolder, string strSubjectLine)
        {
            //
            //  to append the status line of every email to the Alarm log
            //  useage: updateAlarmLogFile(strCurrentGKAfolder, strSubjectLine);

            // create the alarm Log file if missing+

            // Note: this will throw an error if you try to create a file in the C root directory
            // strCurrentGKAfolder must not be @"C:\"
            // it must include a folder or be the D drive @"C:\__SystemLogs\"


            string strFileName = "AlarmLog.txt";
            string alarmLog = strCurrentGKAfolder + strFileName;
            checkFileExistance(strCurrentGKAfolder, strFileName);

            // Get the date/time stamp
            string strNow = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm") + " : ";

            if (strSubjectLine != "Do nothing")
            {
                using StreamWriter writetext = File.AppendText(alarmLog);
                writetext.WriteLine(strNow + strSubjectLine);
                writetext.Close();
            }
        }

        public void updateSystemLogFile(string strLogfolder, string strActivityComment)
        {
            //
            //  to append the status line of every email to the Alarm log
            //  useage: updateAlarmLogFile(strCurrentGKAfolder, strSubjectLine);

            // create the alarm Log file if missing
            // Note: this will throw an error if you try to create a file in the C root directory
            // strCurrentGKAfolder must not be @"C:\"
            // it must include a folder or be the D drive @"C:\__SystemLogs\"


            string strFileName = "SystemActivityLog.txt";
            string activityLog = strLogfolder + strFileName;
            checkFileExistance(strLogfolder, strFileName);



            //if (!Directory.Exists(strLogfolder))
            //{
            //    Directory.CreateDirectory(strLogfolder);
            //}

            //if (!File.Exists(activityLog))
            //{
            //    using StreamWriter writetext = File.CreateText(activityLog);
            //    writetext.WriteLine("Log file created: {0}", DateTime.Now.ToString());
            //    writetext.Close();
            //}

            //Get the date / time stamp

            if (strActivityComment != "Do nothing")
            {
                using StreamWriter writetext = File.AppendText(activityLog);
                writetext.WriteLine(strActivityComment);
                writetext.Close();
            }
        }


        public void checkFileExistance(string strRootFolder, string strFileName)
        {
            string strFolderFile = strRootFolder + strFileName;

            string strTimeNow = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

            if (!Directory.Exists(strRootFolder))
            {
                Directory.CreateDirectory(strRootFolder);
            }

            if (!File.Exists(strFolderFile))
            {
                using StreamWriter writetext = File.CreateText(strFolderFile);
                writetext.WriteLine("File created: {0}", strTimeNow);
                writetext.Close();
            }
            return;

        }

        public string doesFileExist(string strRootFolder, string strFileName)
        {
            string strFolderFile = strRootFolder + strFileName;

            string strAnswer = "Yes";

            if (!File.Exists(strFolderFile))
            {
                strAnswer = "No";
            }
            return strAnswer;

        }


        public string updateNoDataAlarmFile(string strAlarmfolder, string strCurrentAlarmState)
        {
            //
            // it must include a folder or be the D drive @"C:\__SystemAlarms\"

            string strTimeNow = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

            string strFileName = "NoDataAlarmState.txt";
            string alarmLog = strAlarmfolder + strFileName;
            checkFileExistance(strAlarmfolder, strFileName);

            //Read the current alarm state

            string strPreviousAlarmState = File.ReadAllText(alarmLog);
            string strAlarmTime = "";

            if (strPreviousAlarmState.Contains("File created:"))
            {
                strAlarmTime = "2000-01-01 00:01";
                strPreviousAlarmState = "Alarm state not determined";
            }
            else
            {
                string[] AlarmStateDetails = strPreviousAlarmState.Split("/");
                strAlarmTime = AlarmStateDetails[0];
                strPreviousAlarmState = AlarmStateDetails[1];
            }

            string strSMSaction;
            if (strCurrentAlarmState != strPreviousAlarmState)
            {
                strSMSaction = "SendSMS";
            }
            else
            {
                strSMSaction = "DoNotSendSMS";
            }

            strCurrentAlarmState = strTimeNow + "/" + strCurrentAlarmState;
            using StreamWriter txtAction = File.CreateText(alarmLog);
            txtAction.Write(strCurrentAlarmState);
            txtAction.Close();

            return strSMSaction;

        }




        public string updateSystemAlarmFile(string strAlarmfolder, string strCurrentAlarmState)
        {
            //
            // it must include a folder or be the D drive @"C:\__SystemAlarms\"


            string strFileName = "SPN010_AlarmState.txt";
            string alarmLog = strAlarmfolder + strFileName;
            checkFileExistance(strAlarmfolder, strFileName);

            //if (!Directory.Exists(strAlarmfolder))
            //{
            //    Directory.CreateDirectory(strAlarmfolder);
            //}

            //if (!File.Exists(alarmLog))
            //{
            //    using StreamWriter writetext = File.CreateText(alarmLog);
            //    writetext.WriteLine("Empty File");
            //    writetext.Close();
            //}

            //Read the current alarm state

            string strPreviousAlarmState = File.ReadAllText(alarmLog);

            string strSMSaction;
            if (strCurrentAlarmState != strPreviousAlarmState)
            {
                strSMSaction = "SendSMS";
            }
            else
            {
                strSMSaction = "DoNotSendSMS";
            }

            using StreamWriter txtAction = File.CreateText(alarmLog);
            txtAction.Write(strCurrentAlarmState);
            txtAction.Close();

            return strSMSaction;

        }





        //================[ eMail Methods ]=============================================================================

        public void sendEmail(string strEmailLogin, string strEmailPassword, string strEmailFrom, string strEmailRecipients, string strSubjectLine, string strEmailBody)
        {
            //
            // Purpose:
            //      To send a simple email with no attachment
            // Useage:
            //      gna.sendEmail(strEmailLogin, strEmailPassword, strEmailFrom, strEmailRecipients, strSubjectLine,strEmailBody);
            //
            // Information
            //      https://www.emailarchitect.net/easendmail/kb/csharp.aspx?cat=3
            //

            // updated with the 2024 license

            try
            {
                SmtpMail oMail = new("ES-E1646458156-01989-1A16E5275AF48A22-686E917424789B4F")
                {
                    From = strEmailFrom,
                    To = new AddressCollection(strEmailRecipients),
                    Subject = strSubjectLine,
                    TextBody = strEmailBody
                };

                // SMTP server address
                SmtpServer oServer = new("smtp.gmail.com")
                {
                    User = strEmailLogin,
                    Password = strEmailPassword,
                    ConnectType = SmtpConnectType.ConnectTryTLS,
                    Port = 587
                };

                //Set sender email address, please change it to yours
                SmtpClient oSmtp = new();
                oSmtp.SendMail(oServer, oMail);

                Console.WriteLine("email sent...");
            }
            catch (Exception ep)
            {
                Console.WriteLine("Failed to send email:");
                Console.WriteLine("");
                Console.WriteLine(strEmailLogin);
                Console.WriteLine(strEmailPassword);
                Console.WriteLine("\nPossible issues:");
                Console.WriteLine("1. Two step verification must be activated");
                Console.WriteLine("2. Use the correct Apps password");
                Console.WriteLine("3. See the T4D Cookbook for step by step details");
                Console.WriteLine("");
                Console.WriteLine(ep.Message);
                Console.ReadKey();
            }
        }



        //================[ SMS Methods ]============================================================================


        public void sendSMS(string strMessage, string strPhone1, string strPhone2, string strJAGaction)
        {
            string accountSid = "AC9fb2a010982500ddb640c5a108f57a28";
            string authToken = "e3bcb8e6aa257fcd3db33af0e77da99b";
            string strJAGUKphone = "+447729481504";
            string strJAGDEphone = "+4917672997904";
            strMessage = addCopyright("",strMessage);

            // initialise the TwilioClient
            Twilio.TwilioClient.Init(accountSid, authToken);

            if ((strPhone1 != "None") && (strPhone1[..1] == "+"))
            {
                //Console.WriteLine("1st sms");

                Console.WriteLine("     SMS to: " + strPhone1.Trim());
                var message = MessageResource.Create(
                     body: strMessage,
                     from: new Twilio.Types.PhoneNumber("+447782585652"),
                     to: new Twilio.Types.PhoneNumber(strPhone1.Trim())
                     );
            }

            if ((strPhone2 != "None") && (strPhone2[..1] == "+"))
            {
                //Console.WriteLine("2nd sms");

                Console.WriteLine("     SMS to: " + strPhone2.Trim());
                var message = MessageResource.Create(
                    body: strMessage,
                    from: new Twilio.Types.PhoneNumber("+447782585652"),
                    to: new Twilio.Types.PhoneNumber(strPhone2.Trim())
                    );
            }

            if (strJAGaction == "Yes")
            {

                Console.WriteLine("     SMS to: JAG");
                var message = MessageResource.Create(
                    body: strMessage,
                    from: new Twilio.Types.PhoneNumber("+447782585652"),
                    to: new Twilio.Types.PhoneNumber(strJAGDEphone.Trim())
                    );
            }

            if (strJAGaction == "DE")
            {

                Console.WriteLine("     SMS to: JAG");
                var message = MessageResource.Create(
                    body: strMessage,
                    from: new Twilio.Types.PhoneNumber("+447782585652"),
                    to: new Twilio.Types.PhoneNumber(strJAGDEphone.Trim())
                    );
            }

            if (strJAGaction == "UK")
            {

                Console.WriteLine("     SMS to: JAG");
                var message = MessageResource.Create(
                    body: strMessage,
                    from: new Twilio.Types.PhoneNumber("+447782585652"),
                    to: new Twilio.Types.PhoneNumber(strJAGUKphone.Trim())
                    );
            }
        }

        public void sendSMSArray(string strSMSmessage, string[] smsMobile)
        {
            //
            // Purpose:
            //      To send an SMS message to up to 9 recipients
            // Useage:
            //      gna.sendSMS(strSMSmessage);
            //
            try
            {
                string accountSid = "AC9fb2a010982500ddb640c5a108f57a28";
                string authToken = "e3bcb8e6aa257fcd3db33af0e77da99b";
                strSMSmessage = addCopyright("", strSMSmessage);

                // initialise the TwilioClient
                Twilio.TwilioClient.Init(accountSid, authToken);

                for (int i = 1; i <= 9; i++)
                {
                    if (smsMobile[i] != "None")
                    {
                        var message = MessageResource.Create(
                         body: strSMSmessage,
                         from: new Twilio.Types.PhoneNumber("+447782585652"),
                         to: new Twilio.Types.PhoneNumber(smsMobile[i].Trim())
                         );
                    }
                }
            }
            catch (Exception ep)
            {
                Console.WriteLine("Failed to send SMS (gneGeneralTools/sendSMS)");
                Console.WriteLine("");
                Console.WriteLine(ep.Message);
                Console.ReadKey();
            }
        }



    }
}