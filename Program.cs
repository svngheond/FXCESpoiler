using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;

namespace FXCE
{
    class Program
    {
        private static HttpClient httpClient;
        private static string apiParticipantListUrl = "https://arena.fxce.com/api/contests/{0}/trading_accounts?lang=en&per=10&page={1}&direction=desc&sort_field=total_profit";
        private static string apiParticipantDetailUrl = "https://stp-api.fxce.com/api/trading_accounts/{0}/latest?lang=vi";
        private static string apiSignalCopylUrl = "https://stp-api.fxce.com/api/trading_accounts/{0}/trading_signal_provides/providers?lang=vi&page=1&per=100";
        private static string apiSignalOfMasterUrl = "https://stp-api.fxce.com/api/trading_accounts/{0}/related?lang=vi&page=-1";
        private static string contestsUrl = "https://arena.fxce.com/api/contests/{0}?lang=en";
        private static string apiGigaCollectionUrl = "https://ea.fxce.com/api/posts?lang=vi&page=1&per=1000&sort_fields[created_at]=desc&filters[related_categories]=63243723f0fed10001abd866";
        private static string postGigaCollectionUrl = "https://ea.fxce.com/post/{0}_{1}";
        private static string accountUrl = "https://www.fxce.com/trader-detail/{1}/{0}/account";
        private static string FXCEReport;
        private static List<int> lstType = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9 };
        private static void WriteAuthor()
        {
            Console.WriteLine("###############################################################");
            Console.WriteLine("#                    FXCE Spoiler Signal                      #");
            Console.WriteLine("#    Công cụ phân tích, so sánh các tín hiệu trên sàn FXCE    #");
            Console.WriteLine("#             Tác giả: https://t.me/tranthao8899              #");
            Console.WriteLine("###############################################################");
            Console.WriteLine();
        }

        static void Main(string[] args)
        {
            //Tạo thư mục Report nếu chưa có
            bool exists = System.IO.Directory.Exists(Directory.GetCurrentDirectory() + "\\Report\\");
            if (!exists)
                System.IO.Directory.CreateDirectory(Directory.GetCurrentDirectory() + "\\Report\\");

            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            Console.OutputEncoding = Encoding.UTF8;
            WriteAuthor();

        batdau:
            Console.WriteLine("1) Đọc hướng dẫn");
            Console.WriteLine("2) Phân tích chi tiết 1 tín hiệu");
            Console.WriteLine("3) Phân tích các tài khoản tham gia trong 1 cuộc thi Arena");
            Console.WriteLine("4) Phân tích các tín hiệu đang copy của 1 tài khoản");
            Console.WriteLine("5) Phân tích các tín hiệu đang có của 1 tài khoản");
            Console.WriteLine("6) Phân tích các tín hiệu đang chạy forward test theo chương trình Giga Collection");
            Console.WriteLine("7)*Phân tích các tín hiệu có trong 1 bộ lọc riêng của bạn");
            Console.WriteLine("8)*Phân tích các tín hiệu mà bạn đang theo dõi");
            Console.WriteLine("9) Thoát!");
            int type = GetType();

            try
            {
                switch (type)
                {
                    case 1:
                        Console.Clear();
                        WriteAuthor();
                        Process.Start(@"https://toididautu.com/fxce-spoiler-signal");
                        break;
                    case 2:
                        Console.Clear();
                        WriteAuthor();
                        PhanTichChiTietTinHieu();
                        break;
                    case 3:
                        Console.Clear();
                        WriteAuthor();
                        PhanTichCuocThiArena();
                        break;
                    case 4:
                        Console.Clear();
                        WriteAuthor();
                        PhanTichTinHieuCopy();
                        break;
                    case 5:
                        Console.Clear();
                        WriteAuthor();
                        PhanTichTinHieuCuaTaiKhoan();
                        break;
                    case 6:
                        Console.Clear();
                        WriteAuthor();
                        PhanTichTinHieuGigaCollection();
                        break;
                    case 7:
                        Console.Clear();
                        WriteAuthor();
                        //PhanTichTinHieuFilter();
                        Console.WriteLine("Mình đang xây dựng. Mời bạn quay lại sau!");
                        Console.WriteLine("Ấn phím Enter để tiếp tục...");
                        Console.ReadLine();
                        break;
                    case 8:
                        Console.Clear();
                        WriteAuthor();
                        //PhanTichTinHieuFilter();
                        Console.WriteLine("Mình đang xây dựng. Mời bạn quay lại sau!");
                        Console.WriteLine("Ấn phím Enter để tiếp tục...");
                        Console.ReadLine();
                        break;
                    case 9:
                        Console.WriteLine("Xin chào và hẹn gặp lại!");
                        Thread.Sleep(2000);
                        Environment.Exit(0);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Có lỗi: " + ex.ToString());
                Console.WriteLine("Ấn phím Enter để tiếp tục...");
                Console.ReadLine();
            }

            Console.Clear();
            WriteAuthor();
            goto batdau;
        }
        private static string GetCapchaToken()
        {


            string token = string.Empty;

            string url = "https://www.fxce.com/login";
            new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig());

            ChromeOptions options = new ChromeOptions();
            options.AddArguments("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36");

            ChromeDriver driver = new ChromeDriver(options);
            driver.Navigate().GoToUrl(url);
            Thread.Sleep(3000);
            string source = driver.PageSource;
            IWebElement nodeAcc = driver.FindElement(By.XPath("//*[@name='email']"));
            IWebElement nodePass = driver.FindElement(By.XPath("//*[@name='password']"));
            nodeAcc.SendKeys("user");
            nodePass.SendKeys("pass");
            IWebElement nodeLogin = driver.FindElement(By.XPath("//button[@tabindex='1'][contains(text(), 'Login')]"));

            nodeLogin.Click();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((x) =>
            {
                return ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete");
            });
            IWebElement nodeToken = driver.FindElement(By.Id("recaptcha-token"));
            token = nodeToken.GetAttribute("value");
            Console.WriteLine("Token: " + token);
            return token;
        }


        private static void PhanTichChiTietTinHieu()
        {
            FXCEReport = "FXCESpoiler_Detail_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";
            Console.WriteLine("2) Phân tích chi tiết 1 tín hiệu");
            Console.Write("Mời bạn điền URL tài khoản bạn muốn phân tích: ");
            string url = Console.ReadLine();
            Signal signal = new Signal(url);
            string resultSignalCopy = string.Empty;

            string resultUserJson = GetParticipantDetail(signal.Id, signal.Server).Result;
            JObject resultUserObj = JObject.Parse(resultUserJson);
            Console.WriteLine();
            Console.WriteLine("Bắt đầu phân tích tín hiệu copy của tài khoản: " + resultUserObj["data"]["trading_account"]["name"] + " (" + resultUserObj["data"]["trading_account"]["partner_user"]["user"]["username"] + ")");

            Dictionary<string,string> lstSignal = new Dictionary<string, string>();
            lstSignal.Add(signal.Id, signal.Server);
            List<string> lstParticipant = new List<string>();
            lstParticipant.Add("Thứ tự,Tín hiệu,Tài khoản,Số dư (Balance),Equity,P&L,Sụt giảm hiện tại,Tăng trưởng,Tỉ lệ thắng,Sụt giảm lớn nhất,CAGR/MDD,Điểm Fxce,Yếu tố lợi nhuận,Lệnh trung bình/tuần,Thời gian giữ lệnh trung bình,Quỹ đầu tư,Phí copy (FXCE)");
            Console.WriteLine("1. Phân tích: " + resultUserObj["data"]["trading_account"]["name"] + " (" + resultUserObj["data"]["trading_account"]["partner_user"]["user"]["username"] + ")");
            string row = string.Empty;
            row = "1," + resultUserObj["data"]["trading_account"]["name"] + "," + resultUserObj["data"]["trading_account"]["partner_user"]["user"]["username"];
            row += "," + resultUserObj["data"]["trading_account"]["latest_balance"];
            row += "," + resultUserObj["data"]["trading_account"]["latest_equity"];
            double pl = double.Parse(resultUserObj["data"]["trading_account"]["latest_equity"] + "") - double.Parse(resultUserObj["data"]["trading_account"]["latest_balance"] + "");
            row += "," + Math.Round(pl, 2);
            row += "," + Math.Round(pl * 100 / double.Parse(resultUserObj["data"]["trading_account"]["latest_balance"] + ""), 2) + "%";
            row += "," + Math.Round(double.Parse(resultUserObj["data"]["trading_account"]["latest_analysis"]["account_growth"] + ""), 2) + "%";
            row += "," + Math.Round(double.Parse(resultUserObj["data"]["win_rate"] + "") * 100, 2) + "%";
            row += "," + Math.Round(double.Parse(resultUserObj["data"]["trading_account"]["max_equity_drawdown"] + ""), 2) + "%";
            row += "," + Math.Round(double.Parse(resultUserObj["data"]["trading_account"]["recovery_factor"] + ""), 2);
            row += "," + Math.Round(double.Parse(resultUserObj["data"]["trading_account"]["fxce_score"] + ""), 2);
            row += "," + Math.Round(double.Parse(resultUserObj["data"]["trading_account"]["profit_factor"] + ""), 2);
            row += "," + Math.Round(double.Parse(resultUserObj["data"]["avg_trade_per_week"] + ""), 2);
            row += "," + ConvertSecondsToTime(double.Parse(resultUserObj["data"]["avg_trade_length"] + ""));
            row += "," + resultUserObj["data"]["trading_account"]["total_retail_equity"] + "";
            try
            {
                JArray tradingInvestArr = JArray.Parse(resultUserObj["data"]["trading_investment_config"] + "");
                if (tradingInvestArr.Count > 0)
                {
                    JToken tradingInvest = tradingInvestArr.Where(x => x["kind"] + "" == "signal").FirstOrDefault();
                    if (tradingInvest != null)
                        row += "," + tradingInvest["signal_fee"];
                    else
                        row += ",";
                }
                else
                {
                    row += ",";
                }
            }
            catch (Exception)
            {
                row += ",";
            }
            lstParticipant.Add(row);

            SaveToExcel(lstParticipant,lstSignal);
            OpenReport();
        }

        private static void PhanTichTinHieuFilter()
        {
            string capcha = GetCapchaToken();
            string acc = "";
        }
        private static void PhanTichTinHieuGigaCollection()
        {
            FXCEReport = "FXCESpoiler_GigaCollection_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";
            Console.WriteLine("6) Phân tích các tín hiệu đang chạy forward test theo chương trình Giga Collection");
            string postData = GetPostGigaCollections().Result;
            JObject resultObj = JObject.Parse(postData);
            List<string> lstParticipant = new List<string>();
            lstParticipant.Add("Thứ tự,Tín hiệu,Tài khoản,Số dư (Balance),Equity,P&L,Sụt giảm hiện tại,Tăng trưởng,Tỉ lệ thắng,Sụt giảm lớn nhất,CAGR/MDD,Điểm Fxce,Yếu tố lợi nhuận,Lệnh trung bình/tuần,Thời gian giữ lệnh trung bình,Quỹ đầu tư,Phí copy (FXCE)");
            int stt = 1, index = 1;
            JArray postObjs = (JArray)resultObj["data"]["items"];
            Dictionary<string, string> lstSignal = new Dictionary<string, string>();
            foreach (JObject item in postObjs)
            {
                Console.WriteLine(index + ". Phân tích: " + item["title"]);
                string linkForwardTest = GetSignalForwardTest(item["slug"] + "", item["id"] + "");
                if (!string.IsNullOrEmpty(linkForwardTest))
                {
                    Console.WriteLine("  --> Link forward test: " + linkForwardTest);
                    Signal signal = new Signal(linkForwardTest);
                    lstSignal.Add(signal.Id, signal.Server);
                    try
                    {
                        string resultUserJson = GetParticipantDetail(signal.Id, signal.Server).Result;
                        JObject resultUserObj = JObject.Parse(resultUserJson);
                        string row = string.Empty;
                        row = stt + "," + resultUserObj["data"]["trading_account"]["name"] + "," + resultUserObj["data"]["trading_account"]["partner_user"]["user"]["username"];
                        row += "," + resultUserObj["data"]["trading_account"]["latest_balance"];
                        row += "," + resultUserObj["data"]["trading_account"]["latest_equity"];
                        double pl = double.Parse(resultUserObj["data"]["trading_account"]["latest_equity"] + "") - double.Parse(resultUserObj["data"]["trading_account"]["latest_balance"] + "");
                        row += "," + Math.Round(pl, 2);
                        row += "," + Math.Round(pl * 100 / double.Parse(resultUserObj["data"]["trading_account"]["latest_balance"] + ""), 2) + "%";
                        row += "," + Math.Round(double.Parse(resultUserObj["data"]["trading_account"]["latest_analysis"]["account_growth"] + ""), 2) + "%";
                        row += "," + Math.Round(double.Parse(resultUserObj["data"]["win_rate"] + "") * 100, 2) + "%";
                        row += "," + Math.Round(double.Parse(resultUserObj["data"]["trading_account"]["max_equity_drawdown"] + ""), 2) + "%";
                        row += "," + Math.Round(double.Parse(resultUserObj["data"]["trading_account"]["recovery_factor"] + ""), 2);
                        row += "," + Math.Round(double.Parse(resultUserObj["data"]["trading_account"]["fxce_score"] + ""), 2);
                        row += "," + Math.Round(double.Parse(resultUserObj["data"]["trading_account"]["profit_factor"] + ""), 2);
                        row += "," + Math.Round(double.Parse(resultUserObj["data"]["avg_trade_per_week"] + ""), 2);
                        row += "," + ConvertSecondsToTime(double.Parse(resultUserObj["data"]["avg_trade_length"] + ""));
                        row += "," + resultUserObj["data"]["trading_account"]["total_retail_equity"] + "";
                        try
                        {
                            JArray tradingInvestArr = JArray.Parse(resultUserObj["data"]["trading_investment_config"] + "");
                            if (tradingInvestArr.Count > 0)
                            {
                                JToken tradingInvest = tradingInvestArr.Where(x => x["kind"] + "" == "signal").FirstOrDefault();
                                if (tradingInvest != null)
                                    row += "," + tradingInvest["signal_fee"];
                                else
                                    row += ",";
                            }
                            else
                            {
                                row += ",";
                            }
                        }
                        catch (Exception)
                        {
                            row += ",";
                        }

                        lstParticipant.Add(row);
                        stt++;
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("    --> Không thể truy cập vào tài khoản riêng tư!");
                    }
                }
                else
                {
                    Console.WriteLine("  --> Không có link forward test! ");
                }
                index++;
            }
            SaveToExcel(lstParticipant, lstSignal);
            OpenReport();
        }

        private static void PhanTichTinHieuCuaTaiKhoan()
        {
            FXCEReport = "FXCESpoiler_Signal_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";
            Console.WriteLine("5) Phân tích các tín hiệu đang có của 1 tài khoản");
            Console.Write("Mời bạn điền URL tài khoản bạn muốn phân tích các tín hiệu copy của họ: ");
            string url = Console.ReadLine();
            Signal signal = new Signal(url);
            string resultSignalCopy = string.Empty;
            try
            {
                resultSignalCopy = GetSignalOfMaster(signal.Id,signal.Server).Result;
            }
            catch (Exception)
            {
                Console.WriteLine("Có lỗi: Tài khoản đang ở chế độ riêng tư.");
                Console.WriteLine("Ấn phím Enter để tiếp tục...");
                Console.ReadLine();
                return;
            }

            JObject resultSignalCopyObj = JObject.Parse(resultSignalCopy);

            string resultUserJson = GetParticipantDetail(signal.Id, signal.Server).Result;
            JObject resultUserObj = JObject.Parse(resultUserJson);
            Console.WriteLine();
            Console.WriteLine("Bắt đầu phân tích tín hiệu đang có của tài khoản: " + resultUserObj["data"]["trading_account"]["partner_user"]["user"]["username"]);

            List<string> lstParticipant = new List<string>();
            lstParticipant.Add("Thứ tự,Tín hiệu,Tài khoản,Số dư (Balance),Equity,P&L,Sụt giảm hiện tại,Tăng trưởng,Tỉ lệ thắng,Sụt giảm lớn nhất,CAGR/MDD,Điểm Fxce,Yếu tố lợi nhuận,Lệnh trung bình/tuần,Thời gian giữ lệnh trung bình,Quỹ đầu tư,Phí copy (FXCE)");
            int stt = 1;
            JArray participantObjs = (JArray)resultSignalCopyObj["data"]["items"];
            Dictionary<string, string> lstSignal = new Dictionary<string, string>();
            foreach (JObject participant in participantObjs)
            {
                Console.WriteLine(stt + ". Phân tích: " + participant["name"]);

                string participantId = participant["id"] + "";
                string participantTenant = participant["tenant"] + "";
                lstSignal.Add(participantId, participantTenant);
                string resultJson = GetParticipantDetail(participantId, participantTenant).Result;
                JObject resultDetailObj = JObject.Parse(resultJson);
                string row = string.Empty;
                row = stt + "," + resultDetailObj["data"]["trading_account"]["name"] + "," + resultDetailObj["data"]["trading_account"]["partner_user"]["user"]["username"];
                row += "," + resultDetailObj["data"]["trading_account"]["latest_balance"];
                row += "," + resultDetailObj["data"]["trading_account"]["latest_equity"];
                double pl = double.Parse(resultDetailObj["data"]["trading_account"]["latest_equity"] + "") - double.Parse(resultDetailObj["data"]["trading_account"]["latest_balance"] + "");
                row += "," + Math.Round(pl, 2);
                row += "," + Math.Round(pl * 100 / double.Parse(resultDetailObj["data"]["trading_account"]["latest_balance"] + ""), 2) + "%";
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["latest_analysis"]["account_growth"] + ""), 2) + "%";
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["win_rate"] + "") * 100, 2) + "%";
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["max_equity_drawdown"] + ""), 2) + "%";
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["recovery_factor"] + ""), 2);
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["fxce_score"] + ""), 2);
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["profit_factor"] + ""), 2);
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["avg_trade_per_week"] + ""), 2);
                row += "," + ConvertSecondsToTime(double.Parse(resultDetailObj["data"]["avg_trade_length"] + ""));
                row += "," + resultDetailObj["data"]["trading_account"]["total_retail_equity"] + "";
                try
                {
                    JArray tradingInvestArr = JArray.Parse(resultDetailObj["data"]["trading_investment_config"] + "");
                    if (tradingInvestArr.Count > 0)
                    {
                        JToken tradingInvest = tradingInvestArr.Where(x => x["kind"] + "" == "signal").FirstOrDefault();
                        if (tradingInvest != null)
                            row += "," + tradingInvest["signal_fee"];
                        else
                            row += ",";
                    }
                    else
                    {
                        row += ",";
                    }
                }
                catch (Exception)
                {
                    row += ",";
                }

                lstParticipant.Add(row);
                stt++;
            }
            SaveToExcel(lstParticipant, lstSignal);
            OpenReport();
        }
        private static void PhanTichTinHieuCopy()
        {
            FXCEReport = "FXCESpoiler_Copy_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";
            Console.WriteLine("4) Phân tích các tín hiệu đang copy của 1 tài khoản");
            Console.Write("Mời bạn điền URL tài khoản bạn muốn phân tích các tín hiệu copy của họ: ");
            string url = Console.ReadLine();
            Signal signal = new Signal(url);
            string resultSignalCopy = string.Empty;
            try
            {
                resultSignalCopy = GetSignalCopy(signal.Id).Result;
            }
            catch (Exception)
            {
                Console.WriteLine("Có lỗi: Tài khoản đang ở chế độ riêng tư.");
                Console.WriteLine("Ấn phím Enter để tiếp tục...");
                Console.ReadLine();
                return;
            }

            JObject resultSignalCopyObj = JObject.Parse(resultSignalCopy);

            string resultUserJson = GetParticipantDetail(signal.Id, signal.Server).Result;
            JObject resultUserObj = JObject.Parse(resultUserJson);
            Console.WriteLine();
            Console.WriteLine("Bắt đầu phân tích tín hiệu copy của tài khoản: " + resultUserObj["data"]["trading_account"]["name"] + " (" + resultUserObj["data"]["trading_account"]["partner_user"]["user"]["username"] + ")");

            List<string> lstParticipant = new List<string>();
            lstParticipant.Add("Thứ tự,Tín hiệu,Tài khoản,Số dư (Balance),Equity,P&L,Sụt giảm hiện tại,Tăng trưởng,Tỉ lệ thắng,Sụt giảm lớn nhất,CAGR/MDD,Điểm Fxce,Yếu tố lợi nhuận,Lệnh trung bình/tuần,Thời gian giữ lệnh trung bình,Quỹ đầu tư,Phí copy (FXCE)");
            int stt = 1;
            JArray participantObjs = (JArray)resultSignalCopyObj["data"]["items"];
            Dictionary<string, string> lstSignal = new Dictionary<string, string>();
            foreach (JObject participant in participantObjs)
            {
                Console.WriteLine(stt + ". Phân tích: " + participant["signal_trading_account"]["name"]);

                string participantId = participant["signal_trading_account"]["id"] + "";
                string participantTenant = participant["signal_trading_account_tenant"] + "";
                lstSignal.Add(participantId, participantTenant);
                string resultJson = GetParticipantDetail(participantId, participantTenant).Result;
                JObject resultDetailObj = JObject.Parse(resultJson);
                string row = string.Empty;
                row = stt + "," + resultDetailObj["data"]["trading_account"]["name"] + "," + resultDetailObj["data"]["trading_account"]["partner_user"]["user"]["username"];
                row += "," + resultDetailObj["data"]["trading_account"]["latest_balance"];
                row += "," + resultDetailObj["data"]["trading_account"]["latest_equity"];
                double pl = double.Parse(resultDetailObj["data"]["trading_account"]["latest_equity"] + "") - double.Parse(resultDetailObj["data"]["trading_account"]["latest_balance"] + "");
                row += "," + Math.Round(pl, 2);
                row += "," + Math.Round(pl * 100 / double.Parse(resultDetailObj["data"]["trading_account"]["latest_balance"] + ""), 2) + "%";
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["latest_analysis"]["account_growth"] + ""), 2) + "%";
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["win_rate"] + "") * 100, 2) + "%";
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["max_equity_drawdown"] + ""), 2) + "%";
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["recovery_factor"] + ""), 2);
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["fxce_score"] + ""), 2);
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["profit_factor"] + ""), 2);
                row += "," + Math.Round(double.Parse(resultDetailObj["data"]["avg_trade_per_week"] + ""), 2);
                row += "," + ConvertSecondsToTime(double.Parse(resultDetailObj["data"]["avg_trade_length"] + ""));
                row += "," + resultDetailObj["data"]["trading_account"]["total_retail_equity"] + "";
                try
                {
                    JArray tradingInvestArr = JArray.Parse(resultDetailObj["data"]["trading_investment_config"] + "");
                    if (tradingInvestArr.Count > 0)
                    {
                        JToken tradingInvest = tradingInvestArr.Where(x => x["kind"] + "" == "signal").FirstOrDefault();
                        if (tradingInvest != null)
                            row += "," + tradingInvest["signal_fee"];
                        else
                            row += ",";
                    }
                    else
                    {
                        row += ",";
                    }
                }
                catch (Exception)
                {
                    row += ",";
                }

                lstParticipant.Add(row);
                stt++;
            }
            SaveToExcel(lstParticipant, lstSignal);
            OpenReport();
        }

        static void PhanTichCuocThiArena()
        {
            FXCEReport = "FXCEReport_Arena_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";
            Console.WriteLine("3) Phân tích các tài khoản tham gia trong cuộc thi Arena");
            Console.Write("Mời bạn điền URL cuộc thi bạn muốn phân tích: ");
            string url = Console.ReadLine();
            string contestsID = url.Replace("/vi/", "/").Replace("https://arena.fxce.com/contest/", "");
            if (contestsID.Contains("@"))
                contestsID = contestsID.Split('@')[0];

            string resultArena = GetArenaDetail(contestsID).Result;
            JObject resultArenaObj = JObject.Parse(resultArena);
            Console.WriteLine();
            Console.WriteLine("Bắt đầu phân tích cuộc thi: " + resultArenaObj["data"]["contest_content"]["name"]);

            List<string> lstParticipant = new List<string>();
            lstParticipant.Add("Thứ tự,Tên quỹ,Tài khoản,Số dư (Balance),Equity,P&L,Sụt giảm hiện tại,Tăng trưởng,Sụt giảm lớn nhất,CAGR/MDD,Điểm FXCE,Yếu tố lợi nhuận,Lệnh trung bình/tuần,Thời gian giữ lệnh trung bình,Quỹ đầu tư,Phí copy (FXCE),Trạng thái");
            int pageNum = 10;
            int stt = 1;
            Dictionary<string, string> lstSignal = new Dictionary<string, string>();
            for (int i = 1; i <= pageNum; i++)
            {
                string resultJson = GetListParticipant(i, contestsID).Result;
                JObject resultObj = JObject.Parse(resultJson);
                JArray participantObjs = (JArray)resultObj["data"]["items"];
                pageNum = int.Parse(resultObj["data"]["pagination"]["total_pages"] + "");
                foreach (JObject participant in participantObjs)
                {
                    Console.WriteLine(stt + ". Phân tích: " + participant["name"] + " (" + participant["user"]["username"] + ")");
                    lstSignal.Add(participant["id"] + "", participant["tenant"] + "");
                    string row = string.Empty;
                    row = stt + "," + participant["name"] + "," + participant["user"]["username"];
                    try
                    {
                        resultJson = GetParticipantDetail(participant["id"] + "", participant["tenant"] + "").Result;
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("    --> Không thể truy cập vào tài khoản riêng tư!");
                        continue;
                    }

                    JObject resultDetailObj = JObject.Parse(resultJson);
                    row += "," + resultDetailObj["data"]["trading_account"]["latest_balance"];
                    row += "," + resultDetailObj["data"]["trading_account"]["latest_equity"];
                    double pl = double.Parse(resultDetailObj["data"]["trading_account"]["latest_equity"] + "") - double.Parse(resultDetailObj["data"]["trading_account"]["latest_balance"] + "");
                    row += "," + Math.Round(pl, 2);
                    row += "," + Math.Round(pl * 100 / double.Parse(resultDetailObj["data"]["trading_account"]["latest_balance"] + ""), 2) + "%";
                    row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["latest_analysis"]["account_growth"] + ""), 2) + "%";
                    row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["max_equity_drawdown"] + ""), 2) + "%";
                    row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["recovery_factor"] + ""), 2);
                    row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["fxce_score"] + ""), 2);
                    row += "," + Math.Round(double.Parse(resultDetailObj["data"]["trading_account"]["profit_factor"] + ""), 2);
                    row += "," + Math.Round(double.Parse(resultDetailObj["data"]["avg_trade_per_week"] + ""), 2);
                    row += "," + ConvertSecondsToTime(double.Parse(resultDetailObj["data"]["avg_trade_length"] + ""));
                    row += "," + resultDetailObj["data"]["trading_account"]["total_retail_equity"] + "";
                    try
                    {
                        JArray tradingInvestArr = JArray.Parse(resultDetailObj["data"]["trading_investment_config"] + "");
                        if (tradingInvestArr.Count > 0)
                        {
                            JToken tradingInvest = tradingInvestArr.Where(x => x["kind"] + "" == "signal").FirstOrDefault();
                            if (tradingInvest != null)
                                row += "," + tradingInvest["signal_fee"];
                            else
                                row += ",";
                        }
                        else
                        {
                            row += ",";
                        }
                    }
                    catch (Exception)
                    {
                        row += ",";
                    }
                    string is_disqualified = participant["trading_account_colosseum_view"]["is_disqualified"] + "";
                    row += "," + (bool.Parse(is_disqualified) ? "Bị loại" : "Tham gia");
                    lstParticipant.Add(row);
                    stt++;
                }
            }

            SaveToExcel(lstParticipant,lstSignal);
            OpenReport();
        }
        static void SaveToExcel(List<string> lstParticipant,Dictionary<string,string> lstSignal)
        {
            Console.WriteLine("Đang xuất dữ liệu ra excel");
            Application objXL = new Application();
            Workbook objWB = objXL.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet objSHT = objWB.Worksheets[1];
            for (int i = 0, r = lstParticipant.Count; i < r; i++)
            {
                string[] arr = lstParticipant[i].Split(',');
                for (int j = 0, c = arr.Count(); j < c; j++)
                {
                    if (i > 0 && j == 1)
                    {
                        string accUrl = string.Format(accountUrl, lstSignal.ElementAt(i-1).Key, lstSignal.ElementAt(i-1).Value);
                        objSHT.Cells[i + 1, j + 1].Formula = "=HYPERLINK(\"" + accUrl + "\",\"" + arr[j] + "\")";
                    }
                    else
                        objSHT.Cells[i + 1, j + 1].Value = arr[j];
                }
            }
            Range rangeStyles = objSHT.UsedRange;
            rangeStyles.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            Range rangeHeader = objSHT.Rows[1];
            rangeHeader.Font.Bold = true;
            rangeHeader.AutoFilter(1);
            rangeStyles.Columns.AutoFit();

            objWB.SaveAs(Directory.GetCurrentDirectory() + "\\Report\\" + FXCEReport);
            objWB.Close();
            objXL.Quit();
            Console.WriteLine("Đã lập xong báo cáo: " + FXCEReport);
        }
        static void OpenReport()
        {
            Process process = new Process();
            process.StartInfo.FileName = "excel.exe";
            process.StartInfo.Arguments = "\"" + Directory.GetCurrentDirectory() + "\\Report\\" + FXCEReport + "\"";
            process.Start();
            Console.WriteLine("Ấn phím Enter để tiếp tục...");
            Console.ReadLine();
        }
        private static async Task<string> GetSignalCopy(string userId)
        {
            httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("x-api-version", "v1");

            var response = await httpClient.GetAsync(string.Format(apiSignalCopylUrl, userId));
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }

        private static async Task<string> GetSignalOfMaster(string userId, string tenant)
        {
            httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("x-api-version", "v1");
            httpClient.DefaultRequestHeaders.Add("x-api-tenant", tenant);

            var response = await httpClient.GetAsync(string.Format(apiSignalOfMasterUrl, userId));
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }

        private static int GetType()
        {
            int type = 0;
            do
            {
                try
                {
                    Console.Write("Mời bạn nhập lựa chọn: ");
                    type = int.Parse(Console.ReadLine());
                }
                catch (Exception)
                {
                }

            } while (!lstType.Contains(type));
            return type;
        }
        private static string GetSignalForwardTest(string slug, string id)
        {
            string result = string.Empty;
            var web = new HtmlWeb();
            var doc = web.Load(string.Format(postGigaCollectionUrl, slug, id));
            var node = doc.DocumentNode.SelectSingleNode("//*[contains(text(), 'kết quả giao dịch trên FXCE Social Trading Platform tại')]/a");
            if (node != null && node.Attributes["href"] != null)
            {
                result = node.Attributes["href"].Value;
            }
            return result;
        }
        static async Task<string> GetListParticipant(int page, string contestsID)
        {
            httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("x-api-version", "v1");
            var response = await httpClient.GetAsync(string.Format(apiParticipantListUrl, contestsID, page));
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }
        static async Task<string> GetParticipantDetail(string participantId, string tenant)
        {
            httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("x-api-version", "v1");
            httpClient.DefaultRequestHeaders.Add("x-api-tenant", tenant);

            var response = await httpClient.GetAsync(string.Format(apiParticipantDetailUrl, participantId));
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }

        static async Task<string> GetPostGigaCollections()
        {
            httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("x-api-version", "v1");

            var response = await httpClient.GetAsync(string.Format(apiGigaCollectionUrl));
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }

        static async Task<string> GetArenaDetail(string contestsID)
        {
            httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("x-api-version", "v1");

            var response = await httpClient.GetAsync(string.Format(contestsUrl, contestsID));
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }
        public static string ConvertSecondsToTime(double seconds)
        {
            int minutes = (int)(seconds / 60);
            int hours = minutes / 60;
            minutes %= 60;

            return $"{hours} giờ {minutes} phút";
        }
    }
    public class Signal
    {
        public string Id { get; set; }
        public string Server { get; set; }
        public Signal()
        {
        }
        public Signal(string url)
        {
            if (url.Contains("https://www.fxce.com/trader-detail/"))
            {
                string[] arr = url.Replace("https://www.fxce.com/trader-detail/", "").Split('/');
                Server = arr[0];
                Id = arr[1];
            }
            else if (url.Contains("https://share.fxce.com/t/"))
            {
                string[] arr = url.Replace("https://share.fxce.com/t/", "").Split('/');
                Id = arr[1];
                if (arr[0] == "0")
                {
                    Server = "fxce-mt5-live";
                }
                else if (arr[0] == "1")
                {
                    Server = "fxce-mt5-demo";
                }
                else
                {
                    Server = "fxce-mt5-cent";
                }
            }
        }
    }
}
