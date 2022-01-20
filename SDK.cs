using NumeriX.Pro;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Class1
{
    class Program
    {
        static public List<string> inputListMaker(DateTime dt, string addEq = "0", string addIr = "0", string addMf = "0")
        {
            List<string> temp = new List<string>();

            string inputProcess1 = @"Consoleclient.exe /command:runProjection /grid:true /valuationDate:" + dt.ToString("yyyy-MM-dd")
                + " /JobName:Type1_EQIR /jobsetid:"+ dt.ToString("yyMMdd") + "001"
               + " /projection:Type1_VA /ProjSettings:"+"\"setAddEq=" + addEq + "|setAddIr=" + addIr + "|setAddMf=" + addMf
               + "\" /policypertask:8 /DB:\"Data Source=10.10.24.191,1325;Initial Catalog=HanwhaInternal;User Id=hedge;Password=123123;Timeout=0\"" + " /prefix:Hanwha";
            /*
            string inputProcess2 = @"Consoleclient.exe /command:runProejction /valuationDate:" + dt.ToString("yyyy-MM-dd")
               + " /projection:Type2_VA /ProjSettings:setAddEq=" + addEq + "|setAddIr=" + addIr + "|setAddMf=" + addMf
               + " /policypertask:4 /DB:Data Source=10.10.24.191,1325;Initial Catalog=HanwhaInternal;User Id=hedge;Password=123123;Timeout=0" + " /prefix:Hanwha";
            string inputProcess3 = @"Consoleclient.exe /command:runProejction /valuationDate:" + dt.ToString("yyyy-MM-dd")
                           + " /projection:Type3_VA /ProjSettings:setAddEq=" + addEq + "|setAddIr=" + addIr + "|setAddMf=" + addMf
                           + " /policypertask:31 /DB:Data Source=10.10.24.191,1325;Initial Catalog=HanwhaInternal;User Id=hedge;Password=123123;Timeout=0" + " /prefix:Hanwha";
            string inputProcess4 = @"Consoleclient.exe /command:runProejction /valuationDate:" + dt.ToString("yyyy-MM-dd")
                           + " /projection:Type4_VA /ProjSettings:setAddEq=" + addEq + "|setAddIr=" + addIr + "|setAddMf=" + addMf
                           + " /policypertask:33 /DB:Data Source=10.10.24.191,1325;Initial Catalog=HanwhaInternal;User Id=hedge;Password=123123;Timeout=0" + " /prefix:Hanwha";
            string inputProcess5 = @"Consoleclient.exe /command:runProejction /valuationDate:" + dt.ToString("yyyy-MM-dd")
                           + " /projection:Type5_VA /ProjSettings:setAddEq=" + addEq + "|setAddIr=" + addIr + "|setAddMf=" + addMf
                           + " /policypertask:6 /DB:Data Source=10.10.24.191,1325;Initial Catalog=HanwhaInternal;User Id=hedge;Password=123123;Timeout=0" + " /prefix:Hanwha";
            string inputProcess6 = @"Consoleclient.exe /command:runProejction /valuationDate:" + dt.ToString("yyyy-MM-dd")
                           + " /projection:Type6_VA /ProjSettings:setAddEq=" + addEq + "|setAddIr=" + addIr + "|setAddMf=" + addMf
                           + " /policypertask:4 /DB:Data Source=10.10.24.191,1325;Initial Catalog=HanwhaInternal;User Id=hedge;Password=123123;Timeout=0" + " /prefix:Hanwha";
           */
            temp.Add(inputProcess1);
            //temp.Add(inputProcess2);
            //temp.Add(inputProcess3);
            //temp.Add(inputProcess4);
            //temp.Add(inputProcess5);
            //temp.Add(inputProcess6);
            return temp;
        }



        public static void cmdRUN(DateTime dt, string addEq = "0", string addIr = "0", string addMf ="0")
        {
            List<string> inputProcess = inputListMaker(dt, addEq, addIr, addMf);
            System.Diagnostics.ProcessStartInfo proInfo = new System.Diagnostics.ProcessStartInfo();
            System.Diagnostics.Process pro = new System.Diagnostics.Process();
            
            proInfo.FileName = @"cmd.exe";
            proInfo.Arguments = "/C dir";
            proInfo.CreateNoWindow = false; // Window OK
            proInfo.UseShellExecute = false;
            proInfo.RedirectStandardOutput = true;
            proInfo.RedirectStandardInput = true;
            proInfo.RedirectStandardError = true;

            pro.StartInfo = proInfo;
            
            ///for (int i = 0;i<inputProcess.Count;i++)
            //{
                Console.WriteLine(inputProcess[0]);

                pro.Start();
                pro.StandardInput.Write(inputProcess[0]+Environment.NewLine);
             pro.StandardInput.Close();

            pro.WaitForExit();
            string result = pro.StandardOutput.ReadToEnd();
            Console.WriteLine(result);
            //}
            Console.ReadLine();
            pro.Close();

        }


        public static void Main(string[] args)
        {
            IRSwap irs = new IRSwap();
            DateTime today = new DateTime(2022, 01, 03); /* Today ! */
            DateTime todayDur = irs.getAddDate(today, "52M");
   
            DateTime prevDay = irs.getSubDate(today, "1D");
            DateTime prevDur = irs.getAddDate(prevDay, "52M");

            double rIR_today = irs.getDF(today, todayDur);
            double rIR_todaynosp= irs.getDFnoSpread(today, todayDur);
            double rIR_prev = irs.getDF(prevDay, prevDur);
            Console.WriteLine(rIR_today);
            Console.WriteLine(rIR_todaynosp);
            Console.WriteLine(rIR_prev);
            double rMMF_today = irs.getDF(today, irs.getAddDate(today, "75D"));
            double rMMF_prev = irs.getDF(prevDay, irs.getAddDate(prevDay, "75D"));
 
            Console.WriteLine(rMMF_today);
            Console.WriteLine(rMMF_prev);
            string addedIR = string.Format("{0:F10}",Math.Round(rIR_today / rIR_prev - 1.0,10));
            string addedMF = string.Format("{0:F10}", Math.Round(rMMF_today / rMMF_prev - 1.0, 10));
            string addedEQ = string.Format("{0:F10}", Math.Round(irs.getReq(prevDay, today),10));

            Console.WriteLine(today);
            Console.WriteLine(todayDur);
            Console.WriteLine(prevDay);
            Console.WriteLine(prevDur);
            Console.WriteLine(addedIR);
            Console.WriteLine(addedMF);
            Console.WriteLine(addedEQ);

            // Job Execution //

            //ProjSettings:"setADDEQ=|"

           // cmdRUN(today, addedEQ, addedIR, addedMF);

            //irs.getFixing(today);
        }

    }



    class IRSwap
    {
        Application nxApp = new Application();
        ApplicationWarning nxWarning = new ApplicationWarning();
        public void checkFatal(String ID)
        {
            if (nxWarning.HasFatal())
            {
                String tmp = "Loading " + ID + " had fatal messages: " + nxWarning.GetFatal(0);
                throw new NumeriX.Pro.ApplicationException(tmp);
            }
        }



        public void addSwap(Application app, String ID, String container, String convention, String tenor, double rate)
        {
            ApplicationCall request = new ApplicationCall();
            request.AddValue("ID", ID);
            request.AddValue("OBJECT", "Instrument");
            request.AddValue("TYPE", "Swap");
            request.AddValue("CONTAINER", container);
            request.AddValue("CURRENCY", "KRW");
            request.AddValue("CONVENTION", convention);
            request.AddValue("END TENOR", tenor);
            request.AddValue("RATE/DIVIDEND", rate);

            String[] nullHeaders =
                new String[] { "ID", "Local ID", "TIMER", "TIMER CPU", "UPDATED" };

            app.Call(request, nullHeaders, nxWarning);
            checkFatal(ID);
        }
        public void addClass(Application app, String ID, String container)
        {
            ApplicationCall request = new ApplicationCall();
            request.AddValue("ID", ID);
            request.AddValue("OBJECT", "Class");
            request.AddValue("TYPE", "Container");
            request.AddValue("CONTAINER", container);

            String[] nullHeaders = new String[] { "ID", "Local ID", "TIMER", "TIMER CPU", "UPDATED" };

            app.Call(request, nullHeaders, nxWarning);
            checkFatal(ID);
        }
        public void addSettings_calendar(Application app, String ID)
        {
            ApplicationData nxData = new ApplicationData();
            nxData.AddHeader("DATES");
            nxData.AddValue("DATES", new DateTime(1996, 1, 1));
            nxData.AddValue("DATES", new DateTime(2022, 01, 31));
            nxData.AddValue("DATES", new DateTime(2022, 02, 01));
            nxData.AddValue("DATES", new DateTime(2022, 02, 02));
            nxData.AddValue("DATES", new DateTime(2022, 03, 01));
            nxData.AddValue("DATES", new DateTime(2022, 05, 05));
            nxData.AddValue("DATES", new DateTime(2022, 06, 01));
            nxData.AddValue("DATES", new DateTime(2022, 06, 06));
            nxData.AddValue("DATES", new DateTime(2022, 08, 15));
            nxData.AddValue("DATES", new DateTime(2022, 09, 09));
            nxData.AddValue("DATES", new DateTime(2022, 09, 12));
            nxData.AddValue("DATES", new DateTime(2022, 10, 03));
            nxData.AddValue("DATES", new DateTime(2022, 12, 21));
            nxData.AddValue("DATES", new DateTime(2023, 01, 23));
            nxData.AddValue("DATES", new DateTime(2023, 03, 01));
            nxData.AddValue("DATES", new DateTime(2023, 05, 01));
            nxData.AddValue("DATES", new DateTime(2023, 05, 05));
            nxData.AddValue("DATES", new DateTime(2023, 06, 06));
            nxData.AddValue("DATES", new DateTime(2023, 08, 15));
            nxData.AddValue("DATES", new DateTime(2023, 09, 28));
            nxData.AddValue("DATES", new DateTime(2023, 09, 29));
            nxData.AddValue("DATES", new DateTime(2023, 10, 03));
            nxData.AddValue("DATES", new DateTime(2023, 10, 09));
            nxData.AddValue("DATES", new DateTime(2023, 12, 25));
            nxData.AddValue("DATES", new DateTime(2024, 01, 01));
            nxData.AddValue("DATES", new DateTime(2024, 02, 09));
            nxData.AddValue("DATES", new DateTime(2024, 03, 01));
            nxData.AddValue("DATES", new DateTime(2024, 04, 10));
            nxData.AddValue("DATES", new DateTime(2024, 05, 01));
            nxData.AddValue("DATES", new DateTime(2024, 05, 15));
            nxData.AddValue("DATES", new DateTime(2024, 06, 06));
            nxData.AddValue("DATES", new DateTime(2024, 08, 15));
            nxData.AddValue("DATES", new DateTime(2024, 09, 16));
            nxData.AddValue("DATES", new DateTime(2024, 09, 17));
            nxData.AddValue("DATES", new DateTime(2024, 09, 18));
            nxData.AddValue("DATES", new DateTime(2024, 10, 03));
            nxData.AddValue("DATES", new DateTime(2024, 10, 09));
            nxData.AddValue("DATES", new DateTime(2024, 12, 25));
            nxData.AddValue("DATES", new DateTime(2025, 01, 01));
            nxData.AddValue("DATES", new DateTime(2025, 01, 28));
            nxData.AddValue("DATES", new DateTime(2025, 01, 29));
            nxData.AddValue("DATES", new DateTime(2025, 01, 30));
            nxData.AddValue("DATES", new DateTime(2025, 05, 01));
            nxData.AddValue("DATES", new DateTime(2025, 05, 05));
            nxData.AddValue("DATES", new DateTime(2025, 05, 06));
            nxData.AddValue("DATES", new DateTime(2025, 06, 06));
            nxData.AddValue("DATES", new DateTime(2025, 08, 15));
            nxData.AddValue("DATES", new DateTime(2025, 10, 03));
            nxData.AddValue("DATES", new DateTime(2025, 10, 06));
            nxData.AddValue("DATES", new DateTime(2025, 10, 07));
            nxData.AddValue("DATES", new DateTime(2025, 10, 09));
            nxData.AddValue("DATES", new DateTime(2025, 12, 25));
            nxData.AddValue("DATES", new DateTime(2026, 01, 01));
            nxData.AddValue("DATES", new DateTime(2026, 02, 16));
            nxData.AddValue("DATES", new DateTime(2026, 02, 17));
            nxData.AddValue("DATES", new DateTime(2026, 02, 18));
            nxData.AddValue("DATES", new DateTime(2026, 05, 01));
            nxData.AddValue("DATES", new DateTime(2026, 05, 05));
            nxData.AddValue("DATES", new DateTime(2026, 09, 24));
            nxData.AddValue("DATES", new DateTime(2026, 09, 25));
            nxData.AddValue("DATES", new DateTime(2026, 10, 09));
            nxData.AddValue("DATES", new DateTime(2026, 12, 25));
            nxData.AddValue("DATES", new DateTime(2027, 01, 01));
            nxData.AddValue("DATES", new DateTime(2027, 02, 08));
            nxData.AddValue("DATES", new DateTime(2027, 03, 01));
            nxData.AddValue("DATES", new DateTime(2027, 05, 05));
            nxData.AddValue("DATES", new DateTime(2027, 05, 13));
            nxData.AddValue("DATES", new DateTime(2027, 09, 14));
            nxData.AddValue("DATES", new DateTime(2027, 09, 15));
            nxData.AddValue("DATES", new DateTime(2027, 09, 16));
            nxData.AddValue("DATES", new DateTime(2028, 01, 26));
            nxData.AddValue("DATES", new DateTime(2028, 01, 27));
            nxData.AddValue("DATES", new DateTime(2028, 01, 28));
            nxData.AddValue("DATES", new DateTime(2028, 03, 01));
            nxData.AddValue("DATES", new DateTime(2028, 05, 01));
            nxData.AddValue("DATES", new DateTime(2028, 05, 02));
            nxData.AddValue("DATES", new DateTime(2028, 05, 05));
            nxData.AddValue("DATES", new DateTime(2028, 06, 06));
            nxData.AddValue("DATES", new DateTime(2028, 08, 15));
            nxData.AddValue("DATES", new DateTime(2028, 10, 02));
            nxData.AddValue("DATES", new DateTime(2028, 10, 03));
            nxData.AddValue("DATES", new DateTime(2028, 10, 04));
            nxData.AddValue("DATES", new DateTime(2028, 10, 09));
            nxData.AddValue("DATES", new DateTime(2028, 12, 25));
            app.Data(nxData, ID + ".DATES", nxWarning);
            checkFatal(ID);

            ApplicationCall request = new ApplicationCall();
            request.AddValue("ID", ID);
            request.AddValue("OBJECT", "SETTINGS");
            request.AddValue("TYPE", "CALENDAR");
            request.AddValue("CODE", ID);
            request.AddValue("DATES", ID + ".DATES");
            request.AddValue("MONDAY", true);
            request.AddValue("TUESDAY", true);
            request.AddValue("WEDNESDAY", true);
            request.AddValue("THURSDAY", true);
            request.AddValue("FRIDAY", true);
            request.AddValue("SATURDAY", false);
            request.AddValue("SUNDAY", false);

            string[] nullHeaders = new string[] { "ID", "Local ID", "TIMER", "TIMER CPU", "UPDATED" };
            app.Call(request, nullHeaders, nxWarning);
            checkFatal(ID);
        }

        public void addSettings(Application app, String ID)
        {
            ApplicationData nxData = new ApplicationData();

            nxData.AddHeader("NAME");
            nxData.AddHeader("VALUE");

            nxData.AddValue("NAME", "ACCRUAL CALENDAR");
            nxData.AddValue("NAME", "ACCRUAL CONVENTION");
            nxData.AddValue("NAME", "BASIS");
            nxData.AddValue("NAME", "FIX CALENDAR");
            nxData.AddValue("NAME", "FIX CONVENTION");
            nxData.AddValue("NAME", "FREQ");
            nxData.AddValue("NAME", "INTERVAL");
            nxData.AddValue("NAME", "NOTICE PERIOD");
            nxData.AddValue("NAME", "PAY CALENDAR");
            nxData.AddValue("NAME", "PAY CONVENTION");
            nxData.AddValue("NAME", "PRIORITY");
            nxData.AddValue("NAME", "START TENOR");

            nxData.AddValue("VALUE", "SEOUL");
            nxData.AddValue("VALUE", "MF");
            nxData.AddValue("VALUE", "ACT/365");
            nxData.AddValue("VALUE", "SEOUL");
            nxData.AddValue("VALUE", "MF");
            nxData.AddValue("VALUE", "3M");
            nxData.AddValue("VALUE", "1bd");
            nxData.AddValue("VALUE", "1bd");
            nxData.AddValue("VALUE", "SEOUL");
            nxData.AddValue("VALUE", "MF");
            nxData.AddValue("VALUE", 4);
            nxData.AddValue("VALUE", "0bd");

            app.Data(nxData, ID + ".DATA", nxWarning);
            checkFatal(ID);

            ApplicationCall request = new ApplicationCall();
            request.AddValue("ID", ID);
            request.AddValue("OBJECT", "SETTINGS");
            request.AddValue("TYPE", "CONVENTION");
            request.AddValue("DEFAULT VALUES", ID + ".DATA");

            string[] nullHeaders = new string[] { "ID", "Local ID", "TIMER", "TIMER CPU", "UPDATED" };
            app.Call(request, nullHeaders, nxWarning);
            checkFatal(ID);
        }

        public void addInstrumentCollection(Application app, String ID, String container, String instrument)
        {
            ApplicationCall request = new ApplicationCall();
            request.AddValue("ID", ID);
            request.AddValue("OBJECT", "Instrument Collection");
            request.AddValue("TYPE", "Components");
            request.AddValue("CONTAINER", container);
            request.AddValue("CURRENCY", "KRW");
            request.AddValue("INSTRUMENT", instrument);

            String[] nullHeaders = new String[] { "ID", "Local ID", "TIMER", "TIMER CPU", "UPDATED" };
            app.Call(request, nullHeaders, nxWarning);
            checkFatal(ID);
        }


        public void addMarketData(DateTime dt, Application app, String ID, String container, String instruments)
        {
            ApplicationCall request = new ApplicationCall();
            request.AddValue("ID", ID);
            request.AddValue("CONTAINER", container);
            request.AddValue("OBJECT", "Market Data");
            request.AddValue("TYPE", "YIELD");
            request.AddValue("NOWDATE", dt);
            request.AddValue("CURRENCY", "KRW");
            request.AddValue("BASIS", "ACT/365");
            //request.AddValue("EXTRAPOLATION METHOD", "SMITH WILSON");
            //request.AddValue("EXTRAPOLATION PARAMETERS", "SWPARAM");
            request.AddValue("INTERP METHOD", "LOGLINEARDF");
            request.AddValue("INTERP VARIABLE", "DISCOUNT FACTOR");
            request.AddValue("INSTRUMENTS", instruments);

            String[] nullHeaders = new String[] { "ID", "Local ID", "TIMER", "TIMER CPU", "UPDATED" };
            app.Call(request, nullHeaders, nxWarning);
            checkFatal(ID);
        }

        public void addMarketDataSpread(DateTime dt, Application app, String ID, String container)
        {
            ApplicationCall request = new ApplicationCall();
            request.AddValue("ID", ID);
            request.AddValue("CONTAINER", container);
            request.AddValue("OBJECT", "Market Data");
            request.AddValue("TYPE", "YIELD");
            request.AddValue("YIELD", "KRW_IRS");
            request.AddValue("COMMENT","Automatically generated by Leading Hedge");
            request.AddValue("SPREAD", 0.00456);
            request.AddValue("SPREAD BASIS", "ACT/365");
            request.AddValue("COMPOUNDING TYPE", "Annual");
            request.AddValue("NOWDATE", "");
            request.AddValue("CURRENCY", "");
            request.AddValue("BASIS", "");
            request.AddValue("EXTRAPOLATION METHOD", "");
            request.AddValue("EXTRAPOLATION PARAMETERS", "");
            request.AddValue("INTERP METHOD", "");
            request.AddValue("INTERP VARIABLE", "");
            request.AddValue("INSTRUMENTS", "");

            String[] nullHeaders = new String[] { "ID", "Local ID", "TIMER", "TIMER CPU", "UPDATED" };
            app.Call(request, nullHeaders, nxWarning);
            checkFatal(ID);
        }
        public void printDF(DateTime start, DateTime end)
        {
            double printDF = getDF(start, end);
            Console.WriteLine(printDF);
            Console.ReadLine();
        }
        public void printConsle(double number)
        {
            Console.WriteLine(number);
            Console.ReadLine();
        }

        public double getFixing(DateTime dt)
        {
            double result=0;
            double temp = 0;
            string path = @"C:\Program Files\Numerix\Leading Hedge 2.4\client\Hanwha\Database\Market-Data\Fixings\MktFixings.Ir.txt";
            string[] textFile = System.IO.File.ReadAllLines(path);
            

            Dictionary<DateTime, double> dic = new Dictionary <DateTime, double>();
            foreach(var line in textFile)
            {
                //Console.WriteLine(line);
                string[] splitData = line.Split(';');
                if (dic.ContainsKey(Convert.ToDateTime(splitData[2])) == false)
                { 
                    dic.Add(Convert.ToDateTime(splitData[2]), Double.Parse(splitData[3]));
                }
                //Console.WriteLine(Convert.ToDateTime(splitData[2]));
                //Console.WriteLine(Double.Parse(splitData[3]));
                continue;
            }

            // bool hasValue = dic.TryGetValue(dt, out temp);
            if (dic.TryGetValue(dt, out temp))
            {
                //Console.WriteLine(temp);
                result = temp;
            }
            else
            {
                Console.WriteLine("No Fixing!");
            }
            /*
            if (hasValue)
            {
                result = temp;
                Console.WriteLine("Get Fixing !");
            }
            else
            {
                Console.WriteLine("No Fixing !");
            }
            */
            return result;
        }
        
        public double getDF(DateTime start, DateTime end)
        {
            Dictionary<string, double> IRData = getIR(start);

            addClass(nxApp, "GLOBAL", "");
            addClass(nxApp, "Base.Tzero", "GLOBAL");

            addSettings_calendar(nxApp, "SEOUL");
            addSettings(nxApp, "CONVENTIONS::IRSWAP::KRWIRS");

            addSwap(nxApp, "KRW_IRS_KRW.SWAP.3M", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "3M", IRData["3M"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.1Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "1Y", IRData["1Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.2Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "2Y",  IRData["2Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.3Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "3Y",  IRData["3Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.4Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "4Y",  IRData["4Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.5Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "5Y",  IRData["5Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.7Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "7Y",  IRData["7Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.10Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "10Y",IRData["10Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.12Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "12Y",IRData["12Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.15Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "15Y",IRData["15Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.20Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "20Y", IRData["20Y"]);

            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.3M");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.1Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.2Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.3Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.4Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.5Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.7Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.10Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.12Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.15Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.20Y");

            addMarketData(start, nxApp, "KRW_IRS", "GLOBAL^BASE.TZERO", "KRW_IRS.INSTRUM_COLLECTION");
            addMarketDataSpread(start, nxApp, "KRW_IRS.IRS+SPREAD.0.00456", "GLOBAL^BASE.TZERO");
            double df = nxApp.GetDiscountFactor("GLOBAL^BASE.TZERO^KRW_IRS.IRS+SPREAD.0.00456", start, end, nxWarning);
           // double df = nxApp.GetDiscountFactor("GLOBAL^BASE.TZERO^KRW_IRS", start, end, nxWarning);
            return df;
        }


        public double getDFnoSpread(DateTime start, DateTime end)
        {

            Dictionary<string, double> IRData = getIR(start);

            addClass(nxApp, "GLOBAL", "");
            addClass(nxApp, "Base.Tzero", "GLOBAL");

            addSettings_calendar(nxApp, "SEOUL");
            addSettings(nxApp, "CONVENTIONS::IRSWAP::KRWIRS");

            addSwap(nxApp, "KRW_IRS_KRW.SWAP.3M", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "3M", IRData["3M"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.1Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "1Y", IRData["1Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.2Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "2Y", IRData["2Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.3Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "3Y", IRData["3Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.4Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "4Y", IRData["4Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.5Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "5Y", IRData["5Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.7Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "7Y", IRData["7Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.10Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "10Y", IRData["10Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.12Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "12Y", IRData["12Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.15Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "15Y", IRData["15Y"]);
            addSwap(nxApp, "KRW_IRS_KRW.SWAP.20Y", "GLOBAL^BASE.TZERO", "CONVENTIONS::IRSWAP::KRWIRS", "20Y", IRData["20Y"]);

            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.3M");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.1Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.2Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.3Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.4Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.5Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.7Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.10Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.12Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.15Y");
            addInstrumentCollection(nxApp, "KRW_IRS.INSTRUM_COLLECTION", "GLOBAL^BASE.TZERO", "KRW_IRS_KRW.SWAP.20Y");

            addMarketData(start, nxApp, "KRW_IRS", "GLOBAL^BASE.TZERO", "KRW_IRS.INSTRUM_COLLECTION");
            double df = nxApp.GetDiscountFactor("GLOBAL^BASE.TZERO^KRW_IRS", start, end, nxWarning);
            // double df = nxApp.GetDiscountFactor("GLOBAL^BASE.TZERO^KRW_IRS", start, end, nxWarning);
            return df;
        }


        public double getEQ(DateTime today)
        {
            string strDate = today.ToString("yyyyMMdd");
            string path = @"C:\Program Files\Numerix\Leading Hedge 2.4\client\Hanwha\Database\Market-Data\" + strDate + @"_SW\Equity.txt";
            string str = System.IO.File.ReadAllText(path);
            string strResult = str.Split(';')[2];
            double result = System.Convert.ToDouble(strResult);
            return result;
        }
        public Dictionary<string, double> getIR(DateTime today)
        {
            Dictionary<string, double> tempDic= new Dictionary<string, double>();
            string strDate = today.ToString("yyyyMMdd");
            string path = @"C:\Program Files\Numerix\Leading Hedge 2.4\client\Hanwha\Database\Market-Data\" + strDate + @"_SW\YieldCurve_IRS.txt";
            string str = System.IO.File.ReadAllText(path);

            int counter = 0;
            foreach (string line in System.IO.File.ReadLines(path))
            {
                tempDic.Add(line.Split(';')[2], System.Convert.ToDouble(line.Split(';')[6]));
                counter++;
            }
            return tempDic;
        }

        public double getReq(DateTime start, DateTime end)
        {
            double startEQ = getEQ(start);
            double endEQ = getEQ(end);
            double req = endEQ / startEQ-1.0;
            return req;
        }

        public IRSwap()
        {
            addSettings_calendar(nxApp, "SEOUL");
        }

        public DateTime getAddDate(DateTime dt, string advance)
        { 
            DateTime result = nxApp.AddTenor(dt, advance, "MF", "SEOUL", true, nxWarning);
            return result;
        }

        public DateTime getSubDate(DateTime dt, string advance)
        {
            DateTime result = nxApp.SubTenor(dt, advance, "MF", "SEOUL", true, nxWarning);
            return result;
        }
    }

}
