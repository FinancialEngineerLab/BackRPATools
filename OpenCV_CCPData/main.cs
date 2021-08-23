using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using OpenCvSharp;

namespace ConsoleApp4
{
    class Program
    {
        public static string path = @"C:\CCP\"; // main path
        public static double imageSearchTime = 30;
        [DllImport("user32.dll")] static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")] static extern IntPtr FindWindowEx(IntPtr hWnd1, IntPtr hWnd2, string lpsz1, string lpsz2);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr GetWindowTextW(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("User32.dll")] static extern IntPtr SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

        static void Main(string[] args)
        {
            DateTime todaysDate = DateTime.Now;

            GetFile(todaysDate);


        }
        static void GetFile(DateTime todaysDate)
        {
            ProcessStartInfo procInfo = new ProcessStartInfo(@"C:\Program Files\Internet Explorer\iexplore.exe", "https://ccp.krx.co.kr/");
            procInfo.Verb = "runas";
            Process process = Process.Start(procInfo);              

            Bitmap ccp첫화면 = new Bitmap(@"C:\CCP\Pictures\ccp첫화면.png");
            FindImageFromScreen(ccp첫화면);

            IntPtr hWnd = FindWindow(null, "KRX CCP - Internet Explorer");
            SetForegroundWindow(hWnd);

            Thread.Sleep(1000);

            SendKeys.SendWait("{DEL}");
            SendKeys.SendWait("{DEL}");
            SendKeys.SendWait("{DEL}");
            SendKeys.SendWait("{DEL}");
            SendKeys.SendWait("{DEL}");
            SendKeys.SendWait("{DEL}");
            SendKeys.SendWait("{DEL}");
            SendKeys.SendWait("{DEL}");

            SendKeys.SendWait("BM810001");

            SendKeys.SendWait("{TAB}");
            SendKeys.SendWait("hli6363!#");
            SendKeys.SendWait("{Enter}");

            Bitmap 인증서선택 = new Bitmap(@"C:\CCP\Pictures\인증서선택.png");
            FindImageFromScreen(인증서선택);
            Thread.Sleep(3000);
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(100);
            SendKeys.SendWait("{DOWN}");
            Thread.Sleep(100);
            SendKeys.SendWait("{DOWN}");
            Thread.Sleep(100);
            SendKeys.SendWait("{UP}");
            Thread.Sleep(100);
            SendKeys.SendWait("{UP}");
            Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(100);

            SendKeys.SendWait("공인인증서PW");
            Thread.Sleep(100);
            SendKeys.SendWait("{Enter}");

            Thread.Sleep(1000);
            SendKeys.SendWait("{Enter}");

            Thread.Sleep(1000);
            SendKeys.SendWait("{Enter}");

            Bitmap 청산시스템로그인첫화면 = new Bitmap(@"C:\CCP\Pictures\청산시스템로그인첫화면.png");
            OpenCvSharp.Point point = FindImageFromScreen(청산시스템로그인첫화면);
            Thread.Sleep(2000);



            Bitmap 화면번호검색 = new Bitmap(@"C:\CCP\Pictures\화면번호검색.png");
            FindImageClick(화면번호검색);
            Thread.Sleep(200);

            SendKeys.SendWait("PI01");
            SendKeys.SendWait("{Enter}");

            Thread.Sleep(3000);
            Bitmap PI01 = new Bitmap(@"C:\CCP\Pictures\PI01.png");
            FindImageDbClick(PI01);

            Thread.Sleep(1000);

            Bitmap 기준일자 = new Bitmap(@"C:\CCP\Pictures\기준일자.png");
            FindImageClick(기준일자);

            SendKeys.SendWait("{TAB}");
            SendKeys.SendWait("{TAB}");
            SendKeys.SendWait("{TAB}");
            SendKeys.SendWait("{TAB}");
            SendKeys.SendWait("{TAB}");
            SendKeys.SendWait("{TAB}");
            SendKeys.SendWait("{TAB}");
            SendKeys.SendWait("{TAB}");

            SendKeys.SendWait(todaysDate.ToString("yyyyMMdd"));

            Bitmap 조회 = new Bitmap(@"C:\CCP\Pictures\조회.png");
            FindImageClick(조회);

            Thread.Sleep(1000);

            Bitmap 당일거래전송내역 = new Bitmap(@"C:\CCP\Pictures\당일거래전송내역.png");
            FindImageClick(당일거래전송내역);

            Bitmap 채무부담등록포지션 = new Bitmap(@"C:\CCP\Pictures\채무부담등록포지션.png");
            FindImageClick(채무부담등록포지션);

            Thread.Sleep(1000);
            SendKeys.SendWait("{LEFT}");
            SendKeys.SendWait("{LEFT}");
            SendKeys.SendWait("{LEFT}");
            SendKeys.SendWait("{LEFT}");
            SendKeys.SendWait(" ");


            Thread.Sleep(1000);
            SendKeys.SendWait("{PGDN}");

            Bitmap 지정만기별적용금리 = new Bitmap(@"C:\CCP\Pictures\지정만기별적용금리.png");
            FindImageClick(지정만기별적용금리);

            Thread.Sleep(1000);
            SendKeys.SendWait("{LEFT}");
            SendKeys.SendWait("{LEFT}");
            SendKeys.SendWait("{LEFT}");
            SendKeys.SendWait("{LEFT}");
            SendKeys.SendWait(" ");


            Bitmap 단기금리 = new Bitmap(@"C:\CCP\Pictures\단기금리.png");
            FindImageClick(단기금리);
            Thread.Sleep(1000);
            SendKeys.SendWait("{LEFT}");
            SendKeys.SendWait("{LEFT}");
            SendKeys.SendWait("{LEFT}");
            SendKeys.SendWait("{LEFT}");
            SendKeys.SendWait(" ");

            Bitmap 일괄내려받기 = new Bitmap(@"C:\CCP\Pictures\일괄내려받기.png");
            FindImageClick(일괄내려받기);
            Thread.Sleep(500);

            Bitmap 확인 = new Bitmap(@"C:\CCP\Pictures\확인.png");
            FindImageClick(확인);

            Thread.Sleep(10000);
                                                                                                          
            using (SmtpClient client = new SmtpClient { EnableSsl = true, Host = "smtp.gmail.com", Port = 587, Credentials = new NetworkCredential("EMAILID@gmail.com", "EMAILPW") })
            {
                MailAddress from = new MailAddress("EMAILID@gmail.com", "EMAILID", System.Text.Encoding.UTF8);
                MailAddress to = new MailAddress("EMAILRECEIVE@gmail.com");
                MailMessage message = new MailMessage(from, to);

                message.To.Add("EMAILRECEIVE2@gmail.com");
                message.To.Add("EMAILRECEIVE3@gmail.com");

                message.Body = "This e-mail is sent by system.";
                message.BodyEncoding = System.Text.Encoding.UTF8;

                message.Subject = "CCP 데이터";
                message.SubjectEncoding = System.Text.Encoding.UTF8;
                
                // attached files //
                string fPath = @"C:\CCP\" + todaysDate.ToString("yyyyMMdd") + "_00810_TCPMIH20102.txt"; 
                string fPath2 = @"C:\CCP\" + todaysDate.ToString("yyyyMMdd") + "_00810_TCPMIH20202.txt";

                Attachment attachment = new Attachment(fPath);
                Attachment attachment2 = new Attachment(fPath2);

                message.Attachments.Add(attachment);
                message.Attachments.Add(attachment2);
                client.EnableSsl = true;
                client.Send(message);

            }

        }

        static OpenCvSharp.Point FindImageFromScreen(Bitmap find_img)
        {
            DateTime start = DateTime.Now;                              
            DateTime end = DateTime.Now;

            OpenCvSharp.Point findPoint;

            double minval = 0;
            double maxval = 0;
            double spendTime = 0;
            do
            {
                end = DateTime.Now;
                spendTime = (end - start).TotalSeconds;

                Rectangle rect = Screen.AllScreens[0].Bounds;
                Bitmap ScreenBitmap = new Bitmap(rect.Width, rect.Height);

                using (Graphics gr = Graphics.FromImage(ScreenBitmap))
                {
                    gr.CopyFromScreen(rect.Left, rect.Top, 0, 0, rect.Size);
                }

                Mat ScreenMat = OpenCvSharp.Extensions.BitmapConverter.ToMat(ScreenBitmap);
                Mat FindMat = OpenCvSharp.Extensions.BitmapConverter.ToMat(find_img);
                Mat res = ScreenMat.MatchTemplate(FindMat, TemplateMatchModes.CCoeffNormed);

                OpenCvSharp.Point minloc;
                OpenCvSharp.Point maxloc;

                Cv2.MinMaxLoc(res, out minval, out maxval, out minloc, out maxloc);

                findPoint.X = maxloc.X + find_img.Width / 2;
                findPoint.Y = maxloc.Y + find_img.Height / 2;

                ScreenBitmap.Dispose();
                Console.WriteLine("Searching Time (sec) : " + spendTime);

            } while (maxval < 0.9 && spendTime < imageSearchTime);
            if (spendTime < imageSearchTime)
            {
                return findPoint;
            }
            else
            {
                Application.Exit();
                return findPoint;
            }
        }
        static void FindImageClick(Bitmap find_img)
        {
            OpenCvSharp.Point point = FindImageFromScreen(find_img);
            Cursor.Position = new System.Drawing.Point(point.X, point.Y);

            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        }
        static void FindImageDbClick(Bitmap find_img)
        {
            OpenCvSharp.Point point = FindImageFromScreen(find_img);
            Cursor.Position = new System.Drawing.Point(point.X, point.Y);

            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        }

    }
}
