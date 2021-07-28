
using System;

using System.Collections.Generic;

using System.Linq;

using System.Text;

using System.Threading.Tasks;

using System.IO;



namespace CCP

{

    class Program

    {

        static void Main()
        {

            string arg;

            arg = "20210727";



            string txtTCPMIH20102 = @"C:\CCP\" + arg + "_00810_TCPMIH20102.txt";

            string txtTCPMIH20202 = @"C:\CCP\" + arg + "_00810_TCPMIH20202.txt";



            List<string> dataTCPMIH20102 = GetText(txtTCPMIH20102);

            List<string> dataTCPMIH20202 = GetText(txtTCPMIH20202);



            List<string> data = new List<string>();



            for (int i = 0; i < dataTCPMIH20202.Count; i++)

                data.Add(dataTCPMIH20202[i]);



            for (int i = 0; i < dataTCPMIH20102.Count; i++)

                data.Add(dataTCPMIH20102[i]);



            if (!Directory.Exists(@"C:\IRSDATA\" + arg + @"_LQ\"))

            {

                Directory.CreateDirectory(@"C:\IRSDATA\" + arg + @"_LQ\");

            }



            if (File.Exists(@"C:\IRSDATA\" + arg + @"_LQ\YieldCurve_IRS.txt"))

            {

                File.Move(@"C:\IRSDATA\" + arg + @"_LQ\YieldCurve_IRS.txt", @"C:\IRSDATA\" + arg + @"_LQ\YieldCurve_IRS.txt");

            }



            MakeTextFile(data, @"C:\IRSDATA\" + arg + @"_LQ\YieldCurve_IRS.txt");



            DateTime todaysDate = new DateTime(

                Convert.ToInt32(arg.Substring(0, 4))

                , Convert.ToInt32(arg.Substring(4, 2))

                , Convert.ToInt32(arg.Substring(6, 2)));





            // double CD91D = Convert.ToDouble(data[0].Replace("IRSWAP::KRWIRS;SWAP;3M;1d;;;", "").Replace(";0", ""));



            //string str = "KRW;LIBOR;" + todaysDate.ToShortDateString() + "; " + CD91D.ToString();



            // StreamWriter sw = new StreamWriter(@"C:\Program Files\Numerix\Leading Hedge 2.4\client\Hanwha\Database\Market-Data\Fixings\MktFixings.Ir.txt", true);

            //       sw.WriteLine(str);

            //         sw.Close();



        }





        static void MakeTextFile(List<string> data, string opath)

        {

            for (int i = 0; i < data.Count; i++)

                File.WriteAllLines(opath, data);

        }



        static List<string> GetText(string fPath)

        {

            string txt = File.ReadAllText(fPath);



            List<string> data = new List<string>();





            string outputString = "";

            for (int i = 0; i < 16; i++)

            {

                string thisline = txt.Substring(53 + 22 * i, 22);

                thisline = thisline.Replace("  ", " ");

                thisline = thisline.Replace("  ", " ");

                thisline = thisline.Replace("  ", " ");

                thisline = thisline.Replace("  ", " ");

                thisline = thisline.Replace("  ", " ");

                thisline = thisline.Replace(" ", ";");



                string[] split = thisline.Split(';');

                string outstr = "";



                if (split[0] == "1D" || split[0] == "3M" || split[0] == "6M" || split[0] == "9M" || split[0] == "1Y" || split[0] == "18M"

                    || split[0] == "2Y" || split[0] == "3Y" || split[0] == "4Y"

                    || split[0] == "5Y" || split[0] == "6Y" || split[0] == "7Y" || split[0] == "8Y" || split[0] == "9Y" || split[0] == "10Y"

                    || split[0] == "12Y" || split[0] == "15Y" || split[0] == "20Y")

                {

                    outstr =  (Convert.ToDouble(split[1]) / 100.0).ToString();



                    outputString = outputString + outstr + "\n";

                    data.Add(outstr);

                }

            }

            return data;

        }

    }

}
