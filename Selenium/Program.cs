/*
 Packages used:
    OfficeOpenXml
    Selenium
    HtmlAgilityPack
    Chrome
 
 */

using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using OfficeOpenXml;

namespace Selenium
{
    class Program
    {
        static void Main(string[] args)
        {
            IWebDriver driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));

            driver.Navigate().GoToUrl("https://www-banner.aub.edu.lb/pls/weba/bwckschd.p_disp_dyn_sched");
            SelectElement semester = new SelectElement(driver.FindElement(By.Name("p_term")));
            semester.SelectByValue("202010");
            driver.FindElement(By.TagName("form")).Submit();

            #region selecting majors
            SelectElement majors = new SelectElement(driver.FindElement(By.Id("subj_id")));
            //try
            //{
            //    majors.SelectByValue("ACCT");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("AGBU");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("AGSC");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("AHIS");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("AMST");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("ARAB");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("ARCH");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("AROL");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("AVSC");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("BIOC");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("BIOL");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("BIOM");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("BMEN");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("BUSS");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("CHEM");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("CHEN");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("CHIN");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("CIVE");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("CMPS");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("CMTS");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("CVSP");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("DCSN");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("DGRG");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("ECON");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("EDUC");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("EECE");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("ENGL");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("ENHL");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("ENMG");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("ENSC");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("ENST");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("ENTM");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("EPHD");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("EXPR");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("FEAA");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("FINA");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("FREN");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("FSAF");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("FSEC");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("GEOL");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("GRDS");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("HIST");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("HMPD");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("HPCH");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("HUMR");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("IDTH");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("INDE");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("INFO");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("INFP");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("IPEC");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("ISLM");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("LABM");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("LDEM");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MATH");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MAUD");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MBIM");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MCOM");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MECH");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MEST");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MFIN");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MHRM");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MIMG");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MKTG");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MLSP");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MNGT");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MSBA");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MSCU");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("MUSC");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("NFSC");
            //}
            //catch (Exception)
            //{
            //}
            try
            {
                majors.SelectByValue("NURS");
            }
            catch (Exception)
            {
            }
            //try
            //{
            //    majors.SelectByValue("ODFO");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("ORLG");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("PBHL");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("PETR");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("PHIL");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("PHNU");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("PHRM");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("PHYL");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("PHYS");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("PPIA");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("PSPA");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("PSYC");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("PSYT");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("RCOD");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("SART");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("SHRP");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("SOAN");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("STAT");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("THTR");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("URDS");
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    majors.SelectByValue("URPL");
            //}
            //catch (Exception)
            //{
            //}
            #endregion

            driver.FindElement(By.TagName("form")).Submit();


            var html = driver.PageSource;

            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);

            // getting the title - crn - name - section of each course
            var htmlTitleNodes = htmlDoc.DocumentNode.SelectNodes("//th[@class='ddtitle']");

            var htmlDataNodes = htmlDoc.DocumentNode.SelectNodes("//table[@class='datadisplaytable']/tbody");

            List<List<string>> Courses = new List<List<string>>();
            Courses.Add(new List<string> { "Title", "CRN", "Name", "Section", "Term", "Credits", "Time 1", "Days 1", "Place 1", "Date Range 1", "Schedule Type 1", "Instructors 1",
            "Time 2", "Days 2", "Place 2", "Date Range 2", "Schedule Type 2", "Instructors 2", "Time 3", "Days 3", "Place 3", "Date Range 3", "Schedule Type 3", "Instructors 3"});

            var unneededData = new List<string> { "", "-", "Type", "Type\r", "Time", "Time\r", "Days", "Days\r", "Where", "Where\r", "Date Range", "Date Range\r", "Schedule Type",
                "Schedule Type\r", "Instructors", "Instructors\r", "Associated", "Syllabus", "View", "Catalog", "Entry", "(P),", "Term:", "Registration", "Dates:", "Levels:", "Main",
                "Campus", "Credits", "Scheduled", "Meeting", "Times", "Scheduled Meeting Times", "Class", "Class\r", "Return to Previous", "Return to Previous\r", "\r" };

            
            for (int i = 1; i < htmlDataNodes.Count; i++)
            {
                var item = htmlDataNodes[i];
            
                var itemRow = item.SelectNodes("//tr");
                List<string> courseData = new List<string>();
                var details = item.InnerText.Split("\n");

                var terms = item.SelectNodes("//span[contains(text(), 'Associated Term')]");

                var creditsDetails = item.SelectNodes("//td[@class='dddefault']");
                var credits = new List<string>();
                for (int j = 0; j < creditsDetails.Count; j++)
                {
                    foreach (var credit in creditsDetails[j].InnerText.Split("Credits"))
                    {
                        var ddetails = credit.Split();
                        if (ddetails.Length >= 40 && ddetails[ddetails.Length - 2] != "")
                        {
                            credits.Add(ddetails[ddetails.Length - 2]);
                        }
                    }
                }

                if (i - 1 < htmlTitleNodes.Count) // adding the course name, title, crn and section
                    courseData.AddRange(htmlTitleNodes[i - 1].InnerText.Split(" - "));

                if (i - 1 < terms.Count) // Adding the term of the course
                    courseData.Add(terms[i - 1].NextSibling.InnerText);

                if (i - 1 < credits.Count) // Adding the number of credits of the course
                    courseData.Add(credits[i - 1]);

                foreach (var detail in details) // adding other course details
                {
                    if (!unneededData.Contains(detail))
                        courseData.Add(detail);
                }
                if (courseData.Count != 0)
                    Courses.Add(courseData);

            }

            for (int i = 0; i < Courses.Count; i++)
            {
                for (int j = 0; j < Courses[i].Count; j++)
                {
                    string temp = Courses[i][j];
                    WriteToCell(i, j, Courses[i][j]);
                }
            }
            Save();
            driver.Close();
        }
        static string path = @"C:\Users\student\Desktop\test1.xlsx";
        static FileInfo file = new FileInfo(path);
        static ExcelPackage package = new ExcelPackage(file);

        public static void WriteToCell(int i, int j, string s)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets["sheet1"];
            worksheet.Cells[i + 1, j + 1].Value = s;
        }

        public static void Save()
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets["sheet1"];
            for (int i = 1; i < 25; i++)
            {
                worksheet.Column(i).Width = worksheet.DefaultColWidth;
            }
            package.Save();
        }
    }
}
