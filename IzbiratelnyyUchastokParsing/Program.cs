using System.Diagnostics;
using System.Drawing;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools;
using OpenQA.Selenium.Support.UI;

namespace IzbiratelnyyUchastokParsing;

public class Program
{
    private static List<string> Addresses = new List<string>();
    private static List<string> DistrictsOld = new List<string>();
    private static List<string> DistrictsNew = new List<string>();
    private static List<string> UidNew = new List<string>();


    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        GetAllDistrictsFromSource();
        GetAllAddressesFromSource();
        Parsing();
    }

    public static void GetAllAddressesFromSource()
    {
        //string ExcelPath = Environment.CurrentDirectory + "\\Address.xlsx";

        string a = "";
        string b = " ";
        string c = null;

        
        string ExcelPath = Environment.CurrentDirectory + "\\Копия Актуальный список сотрудников РВДК 2024.xlsx";
        FileInfo info = new FileInfo(ExcelPath);

        using (ExcelPackage excel = new ExcelPackage(info))
        {
            var worksheet = excel.Workbook.Worksheets["Лист1"];
            try
            {
                for (int i = 0; !string.IsNullOrEmpty(worksheet.Cells[i + 2, 1].Text); i++)
                {
                    string res = "Город " + (worksheet.Cells[i + 2 + DistrictsOld.Count, 12].Text == "" ? "error" : worksheet.Cells[i + 2 + DistrictsOld.Count, 12].Text) + ", " +
                                 (worksheet.Cells[i + 2 + DistrictsOld.Count, 16].Text == "" ? "error" : worksheet.Cells[i + 2 + DistrictsOld.Count(), 16].Text) + ", " +
                                 worksheet.Cells[i + 2 + DistrictsOld.Count(), 18].Text;
                    Addresses.Add(res);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
        }
    }

    public static void GetAllDistrictsFromSource()
    {
        string ExcelPath = Environment.CurrentDirectory + "\\Копия Актуальный список сотрудников РВДК 2024.xlsx";
        FileInfo info = new FileInfo(ExcelPath);

        using (ExcelPackage excel = new ExcelPackage(info))
        {
            var worksheet = excel.Workbook.Worksheets["Лист1"];
            try
            {
                for (int i = 0; !string.IsNullOrEmpty(worksheet.Cells[i + 2, 1].Text); i++)
                {
                    string res = worksheet.Cells[i + 2, 21].Text;
                    if (res != "")
                    {
                        DistrictsOld.Add(res);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                //System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
        }
    }

    public static void Parsing()
    {
        try
        {
            List<string> test = new List<string>();
            var options = new ChromeOptions();
            options.AddArgument("--no-sandbox");
            options.AddArgument("--start-maximized");
            //options.AddArgument("--headless");     
            //options.AddArgument("--ignore-certificate-errors");
            options.AddArgument("--disable-popup-blocking");
            options.AddArgument("--incognito");
            options.AddUserProfilePreference("safebrowsing.enabled", true);
            for (int i = 0; i < Addresses.Count; i++)
            {
                using (IWebDriver driver = new ChromeDriver(@"C:\VS project\IzbiratelnyyUchastokParsing\IzbiratelnyyUchastokParsing\bin\Debug\net7.0\selenium-manager", options,
                           TimeSpan.FromMinutes(5)))
                {
                    var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(340));
                    driver.Navigate().GoToUrl(@"http://www.cikrf.ru/digital-services/naydi-svoy-izbiratelnyy-uchastok/");
                    bool banner = driver.FindElements(By.XPath(@"/html/body/div[2]/div/div/button")).Count > 0;
                    if (banner)
                    {
                        driver.FindElement(By.XPath(@"/html/body/div[2]/div/div/button")).Click();
                    }

                    bool search = driver.FindElements(By.XPath(@"/html/body/div[1]/div/div[2]/div/div[4]/form/div[2]/div/div/span/span[1]/span/input")).Count > 0;
                    driver.FindElement(By.XPath(@"/html/body/div[1]/div/div[2]/div/div[4]/form/div[2]/div/div/span/span[1]/span/input")).SendKeys(Addresses[i]);
                    Thread.Sleep(1000);
                    bool list = driver.FindElements(By.XPath(@"//*[@id='select2-id2-results']")).Count > 0;
                    var List_error = driver.FindElements(By.XPath(@"//*[@id='select2-id2-results']"));
                    var List_result = driver.FindElement(By.XPath(@"/html/body/span/span/span[2]/ul")).FindElements(By.TagName("li"));
                    if (List_error.Count() == 1 && List_error[0].Text == "Совпадений не найдено")
                    {
                        DistrictsNew.Add("Район не найден");
                        UidNew.Add("Uid не найден");
                        WriteExcel_New();
                        driver.Close();
                        driver.Quit();
                        driver.Dispose();
                    }
                    else if (List_result.Count() == 1)
                    {
                        driver.FindElement(By.XPath(@"//*[@id='select2-id2-results']/li[1]")).Click();
                        Thread.Sleep(2000);
                        driver.FindElement(By.XPath(@"//*[@id='send']")).Click();
                        //Thread.Sleep(2500);
                        wait.Until(d => d.FindElements(By.XPath(@"/html/body/div[1]/div/div[2]/div/div[4]/div[1]/div[2]/div[1]/div[1]/span")).Count > 0);
                        var itogo = driver.FindElement(By.XPath(@"/html/body/div[1]/div/div[2]/div/div[4]/div[1]/div[2]/div[1]/div[1]/span")).Text;

                        Regex regex1 = new Regex("(Ворошиловский)|(Железнодорожный)|(Кировский)|(Ленинский)|(Октябрьский)|(Первомайский)|(Пролетарский)|(Советский)");
                        MatchCollection districts = regex1.Matches(itogo);

                        Regex regex = new Regex("(№[0-9]+)");
                        MatchCollection uid = regex.Matches(itogo);

                        DistrictsNew.Add(districts.Count > 0 ? districts[0].Value.ToString() : "Район не найден");
                        UidNew.Add(uid.Count > 0 ? uid[0].Value.ToString() : "Uid не найден");
                        WriteExcel_New();
                        driver.Close();
                        driver.Quit();
                        driver.Dispose();
                    }
                    else
                    {
                        driver.FindElement(By.XPath(@"//*[@id='select2-id2-results']/li[2]")).Click();
                        Thread.Sleep(1000);
                        bool list1 = driver.FindElements(By.XPath(@"//*[@id='select2-id2-results']")).Count > 0;
                        if (list1)
                        {
                            driver.FindElement(By.XPath(@"//*[@id='select2-id2-results']/li[2]")).Click();
                        }
                        bool list2 = driver.FindElements(By.XPath(@"//*[@id='select2-id2-results']")).Count > 0;
                        if (list2)
                        {
                            driver.FindElement(By.XPath(@"//*[@id='select2-id2-results']/li[2]")).Click();
                        }

                        bool list3 = driver.FindElements(By.XPath(@"//*[@id='select2-id2-results']")).Count > 0;
                        if (list3)
                        {
                            driver.FindElement(By.XPath(@"//*[@id='select2-id2-results']/li[2]")).Click();
                        }

                        driver.FindElement(By.XPath(@"//*[@id='send']")).Click();
                        Thread.Sleep(1000);
                        wait.Until(d => d.FindElements(By.XPath(@"/html/body/div[1]/div/div[2]/div/div[4]/div[1]/div[2]/div[1]/div[1]/span")).Count > 0);
                        var AreaText = driver.FindElement(By.XPath(@"/html/body/div[1]/div/div[2]/div/div[4]/div[1]/div[2]/div[1]/div[1]/span")).Text;
                        var UidText = driver.FindElement(By.XPath(@"/html/body/div[1]/div/div[2]/div/div[4]/div[1]/div[1]/h4")).Text;

                        Regex regex1 = new Regex("(Ворошиловский)|(Железнодорожный)|(Кировский)|(Ленинский)|(Октябрьский)|(Первомайский)|(Пролетарский)|(Советский)");
                        MatchCollection districts = regex1.Matches(AreaText);

                        Regex regex2 = new Regex("(№[0-9]+)");
                        MatchCollection uid = regex2.Matches(UidText);

                        DistrictsNew.Add(districts.Count > 0 ? districts[0].Value.ToString() : "Район не найден");
                        UidNew.Add(uid.Count > 0 ? uid[0].Value.ToString() : "Uid не найден");
                        WriteExcel_New();
                        driver.Close();
                        driver.Quit();
                        driver.Dispose();
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
            throw;
        }
        finally
        {
            foreach (var proc in Process.GetProcessesByName("IEDriverServer"))
            {
                proc.Kill();
            }
        }
    }

    #region Запись в excel новые адреса и уид

    public static void WriteExcel_New()
    {
        string ExcelPath = Environment.CurrentDirectory + "\\Копия Актуальный список сотрудников РВДК 2024.xlsx";
        FileInfo info = new FileInfo(ExcelPath);
        if (info.Exists)
        {
            using (ExcelPackage excel = new ExcelPackage(info))
            {
                for (int i = 0; i < DistrictsNew.Count; i++)
                {
                    var worksheet = excel.Workbook.Worksheets["Лист1"];
                    worksheet.Cells[i + 2 + DistrictsOld.Count(), 21].Value = DistrictsNew[i];
                    worksheet.Cells[i + 2 + DistrictsOld.Count(), 22].Value = UidNew[i];
                    
                }
                excel.Save();
            }
        }
    }

    #endregion
}