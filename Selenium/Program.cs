/*
Author: MOHAMMED EZZEDINE

Description: Parsing Html page to extract data table containing the courses of the American University of Beirut (AUB) during a certain semester.
    Target Pages: https://www-banner.aub.edu.lb/pls/weba/bwckschd.p_disp_dyn_sched
                  https://www-banner.aub.edu.lb/catalog/schd_ + {every capital letter in the alphabet} + .htm

Packages used:
    >   EPPlus.Core
    >   HtmlAgilityPack.NetCore
    >   Selenium.WebDriver
    >   Selenium.WebDriver.ChromeDriver
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
using System.Diagnostics;
using OpenQA.Selenium.Interactions;

namespace Selenium
{
    class Program
    {
        static void Main(string[] args)
        {
            ExtractCourses("202010");
        }

        static void ExtractCourses(string semesterNumber)
        {
            Stopwatch watch = new Stopwatch();
            watch.Start();

            IWebDriver driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));

            // 1st step: Get the crn, actual enrollment, remaining seats, and linked crns for EVERY course under EVERY letter in the dynmaic schedule in banner
            string[] alphabets = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            List<string[]> dynamicCourseDetails = new List<string[]>();

            for (int v = 0; v < 26; v++)
            {
                driver.Navigate().GoToUrl("https://www-banner.aub.edu.lb/catalog/schd_" + alphabets[v] + ".htm");

                var htmlDynam = driver.PageSource;

                var htmlDocDynam = new HtmlDocument();
                htmlDocDynam.LoadHtml(htmlDynam);

                var dynamicRows = htmlDocDynam.DocumentNode.SelectNodes("//p/table/tbody/tr");

                for (int u = 0; u < dynamicRows.Count; u++)
                {
                    var rowDetails = dynamicRows[u].ChildNodes/*.SelectNodes("//td")*/;

                    if (rowDetails.Count == 71 && rowDetails[0].InnerText.ToString().Contains(semesterNumber))
                        // 2: crn; 18: actuall enrollment; 20: available; 70: linked crns
                        dynamicCourseDetails.Add(new string[]
                            {
                                rowDetails[2].InnerText.ToString(),
                                rowDetails[18].InnerText.ToString(),
                                rowDetails[20].InnerText.ToString(),
                                rowDetails[70].InnerText.ToString()
                            });
                }
            }

            // 2nd step extract the other Courses information of each course from the link below
            driver.Navigate().GoToUrl("https://www-banner.aub.edu.lb/pls/weba/bwckschd.p_disp_dyn_sched");
            SelectElement semester;

            string[] possibilities = new string[] { "ACCT", "AGBU", "AGSC", "AHIS", "AMST", "ARAB", "ARCH", "AROL", "AVSC",  "BIOC", "BIOL", "BIOM", "BMEN", "BUSS",
                "CHEM", "CHEN", "CHIN", "CIVE", "CMPS", "CMTS", "CVSP", "DCSN", "DGRG", "ECON", "EDUC", "EECE", "ENGL", "ENHL", "ENMG", "ENSC", "ENST", "ENTM", "EPHD",
                "EXCH", "EXPR", "FEAA", "FINA", "FREN", "FSAF", "FSEC", "GEOL", "GRDS", "HIST", "HMPD", "HPCH", "HUMR", "IDTH", "INDE", "INFO", "INFP", "IPEC",
                "ISLM", "LABM", "LDEM", "MATH", "MAUD", "MBIM", "MCOM", "MECH", "MEST", "MFIN", "MHRM", "MIMG", "MKTG", "MLSP", "MSCU", "MNGT", "MSBA", "MUSC",
                "NFSC", "NURS", "ODFO", "ORLG", "PBHL", "PETR", "PHIL", "PHNU", "PHRM", "PHYL", "PHYS", "PPIA", "PSPA", "PSYC", "PSYT", "RCOD", "SART", "SHRP",
                "SOAN", "STAT", "THTR", "URDS", "URPL"};

            List<List<string>> Courses = new List<List<string>>();

            semester = new SelectElement(driver.FindElement(By.Name("p_term")));
            semester.SelectByValue(semesterNumber);
            driver.FindElement(By.TagName("form")).Submit();
            SelectElement majors = new SelectElement(driver.FindElement(By.Id("subj_id")));

            // select all the majors
            for (int u = 0; u < possibilities.Length; u++)
            {
                string poss = possibilities[u];
                try
                {
                    majors.SelectByValue(poss);
                }
                catch (Exception)
                {
                }
            }
            driver.FindElement(By.TagName("form")).Submit();

            var html = driver.PageSource;

            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);

            // getting the title - crn - name - section of each course
            var htmlTitleNodes = htmlDoc.DocumentNode.SelectNodes("//th[@class='ddtitle']");

            // getting the whole div that contains the data of each course
            var htmlDataNodes = htmlDoc.DocumentNode.SelectNodes("//table[@class='datadisplaytable']/tbody");

            Courses.Add(new List<string> { "Title", "CRN", "Name", "Section", "Term", "Credits", "Actuall Enrollment", "Remaining Seats", "Linked Crns",
                "Time 1", "Days 1", "Place 1", "Schedule Type 1", "Instructors 1", "Time 2", "Days 2", "Place 2", "Schedule Type 2", "Instructors 2",
                "Time 3", "Days 3", "Place 3", "Schedule Type 3", "Instructors 3", "Time 4", "Days 4", "Place 4", "Schedule Type 4", "Instructors 4",
                "Time 5", "Days 5", "Place 5", "Schedule Type 5", "Instructors 5", "Time 6", "Days 6", "Place 6", "Schedule Type 6", "Instructors 6",
                "Time 7", "Days 7", "Place 7", "Schedule Type 7", "Instructors 7", "Time 8", "Days 8", "Place 8", "Schedule Type 8", "Instructors 8",
                "Time 9", "Days 9", "Place 9", "Schedule Type 9", "Instructors 9"});


            if (htmlDataNodes != null)
            {
                var dataNodes = htmlDataNodes[0].InnerText.Split("\r\n\r\n\r\n\r\n\r\n\r\n\r\n\r\n");

                for (int i = 0; i < dataNodes.Length; i++)
                {
                    var item = dataNodes[i].Replace("Type", "").Replace("Time", "").Replace("Days", "").Replace("Where", "").Replace("Date Range", "").Replace("Schedule Type", "")
                        .Replace("Intructors", "").Replace("Scheduled Meeting Times", "").Replace("View Catalog Entry", "").Replace("Main Campus Campus", "").Replace("Course Syllabus", "");
                    var itemLines = item.Split("\r\n");

                    List<string> courseData = new List<string>();
                    int start;
                    if (itemLines[0].Split(" - ").Length >= 2 && (itemLines[0].Split(" - ")[1] == "13124"))
                        Console.WriteLine("s");
                    if (itemLines[0] != "")
                    {
                        courseData.AddRange(itemLines[0].Split(" - "));
                        start = 1;
                    }
                    else
                    {
                        courseData.AddRange(itemLines[1].Split(" - "));
                        start = 2;
                    }

                    for (int k = start; k < itemLines.Length; k++)
                    {
                        if (itemLines[k].Contains("Associated Term:"))
                            courseData.Add(itemLines[k].Replace("Associated Term: ", ""));
                        else if (itemLines[k].Contains("Credits"))
                        {
                            courseData.Add(itemLines[k].Replace("Credits", ""));
                            if (dynamicCourseDetails[k][0].ToString().Equals(courseData[1].ToString()))
                            {
                                courseData.Add(dynamicCourseDetails[k][1] != "" ? dynamicCourseDetails[k][1] : ""); // Actual Enrollment
                                courseData.Add(dynamicCourseDetails[k][2] != "" ? dynamicCourseDetails[k][2] : ""); // Remaining Seats
                                courseData.Add(dynamicCourseDetails[k][3] != "" ? dynamicCourseDetails[k][3] : ""); // Linked crns
                            }
                            else
                            {
                                bool Found = false;
                                for (int w = 0; w < dynamicCourseDetails.Count; w++)
                                {
                                    if (dynamicCourseDetails[w][0].ToString().Equals(courseData[1].ToString()))
                                    {
                                        Found = true;
                                        courseData.Add(dynamicCourseDetails[w][1] != "" ? dynamicCourseDetails[k][1] : ""); // Actual Enrollment
                                        courseData.Add(dynamicCourseDetails[w][2] != "" ? dynamicCourseDetails[k][2] : ""); // Remaining Seats
                                        courseData.Add(dynamicCourseDetails[w][3] != "" ? dynamicCourseDetails[k][3] : ""); // Linked crns
                                        break;
                                    }
                                }
                                if (!Found)
                                {
                                    courseData.Add("");
                                    courseData.Add("");
                                    courseData.Add("");
                                }
                            }
                        }
                        else if ((itemLines[k] == "Class" || itemLines[k].Contains("Class")))
                        {
                            if (k + 6 < itemLines.Length)
                            {
                                try
                                {
                                    courseData.Add(itemLines[k + 1]);
                                    courseData.Add(itemLines[k + 2]);
                                    courseData.Add(itemLines[k + 3]);
                                    courseData.Add(itemLines[k + 5]);
                                    courseData.Add(itemLines[k + 6].Replace("(P)", ""));
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine(itemLines[0]);
                                }
                            }
                        }
                    }
                    if (itemLines.Length <= 20)
                    {
                        courseData.Add("");
                        courseData.Add("");
                        courseData.Add("");
                        courseData.Add("");
                        courseData.Add("");
                        courseData.Add("");
                    }
                    Courses.Add(courseData);
                }
            }
            //driver.Navigate().Back();

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
            watch.Stop();
            Console.WriteLine(watch.Elapsed); // Approx. Execution Time: 00:00:57.4817500
        }

        static string path = @"C:\Users\student\Desktop\test1.xlsx";
        static FileInfo file = new FileInfo(path);
        static ExcelPackage package = new ExcelPackage(file);

        public static void WriteToCell(int i, int j, string s)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets["sheet1"];
            worksheet.Cells[i + 1, j + 2].Value = s;
        }

        public static void Save()
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets["sheet1"];
            worksheet.Cells.AutoFitColumns(10, 100);
            package.Save();
        }
    }
}
