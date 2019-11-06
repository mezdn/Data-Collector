/*
Author: MOHAMMED EZZEDINE
Email: mae117@mail.aub.edu

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

namespace Selenium
{
    class Program
    {
        static void Main(string[] args)
        {
            // the argument passed represent the semester, e.g. Fall 2019-2020: first the year: 2020, then the number of the semester * 10: Fall = 10, Spring = 20, Summer = 30 => 202010
            ExtractCourses("202010");
        }

        static void ExtractCourses(string semesterNumber)
        {
            // timer to keep track on the duration the code needs to execute completely
            Stopwatch watch = new Stopwatch();
            watch.Start();

            // Chrome driver that we will use to reach the need links
            IWebDriver driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));

            // 1st step: Get the crn, actual enrollment, remaining seats, and linked crns for EVERY course under EVERY letter in the dynmaic schedule in banner
            string[] alphabets = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

            // declaring a new list of arrays of strings, dynamicCourseDetails, to save the crn, actual enrollment, remaining seats, and linked crns for every course as an array of strings,
            // to obtain a list of courses, i.e. list of arrays of strings
            List<string[]> dynamicCourseDetails = new List<string[]>();

            for (int v = 0; v < 26; v++)
            {
                // through driver, navigating to each page in dynamic schedule
                driver.Navigate().GoToUrl("https://www-banner.aub.edu.lb/catalog/schd_" + alphabets[v] + ".htm");

                // extracting the html source code of the page
                var htmlDynam = driver.PageSource;

                var htmlDocDynam = new HtmlDocument();
                htmlDocDynam.LoadHtml(htmlDynam);

                // from the html code, we select the divisions that we need
                var dynamicRows = htmlDocDynam.DocumentNode.SelectNodes("//p/table/tbody/tr");

                for (int u = 0; u < dynamicRows.Count; u++)
                {
                    var rowDetails = dynamicRows[u].ChildNodes;

                    if (rowDetails.Count == 71 && rowDetails[0].InnerText.ToString().Contains(semesterNumber))
                        // add to the dynamicCourseDetails a new course (string[]) containing its crn (at index 2), actuall enrollment (at index 18),
                        // available seats (at index 20), and its linked crns (at index 70)
                        dynamicCourseDetails.Add(new string[]
                        {
                            rowDetails[2].InnerText.ToString(), // crn
                            rowDetails[18].InnerText.ToString(), // actuall enrollment
                            rowDetails[20].InnerText.ToString(), // available seats
                            rowDetails[70].InnerText.ToString() // linked crns
                        });
                }
            }
            // Now we have dynamicCourseDetails which is a list of string[], where each string[] stores crn, actual enrollment, remaining seats, and linked crns for a certain course.

            // 2nd step: extract the other Courses information of each course from the link below
            driver.Navigate().GoToUrl("https://www-banner.aub.edu.lb/pls/weba/bwckschd.p_disp_dyn_sched");
            SelectElement semester;

            // all the courses categories (majors):
            string[] possibilities = new string[] { "ACCT", "AGBU", "AGSC", "AHIS", "AMST", "ARAB", "ARCH", "AROL", "AVSC",  "BIOC", "BIOL", "BIOM", "BMEN", "BUSS",
                "CHEM", "CHEN", "CHIN", "CIVE", "CMPS", "CMTS", "CVSP", "DCSN", "DGRG", "ECON", "EDUC", "EECE", "ENGL", "ENHL", "ENMG", "ENSC", "ENST", "ENTM", "EPHD",
                "EXCH", "EXPR", "FEAA", "FINA", "FREN", "FSAF", "FSEC", "GEOL", "GRDS", "HIST", "HMPD", "HPCH", "HUMR", "IDTH", "INDE", "INFO", "INFP", "IPEC",
                "ISLM", "LABM", "LDEM", "MATH", "MAUD", "MBIM", "MCOM", "MECH", "MEST", "MFIN", "MHRM", "MIMG", "MKTG", "MLSP", "MSCU", "MNGT", "MSBA", "MUSC",
                "NFSC", "NURS", "ODFO", "ORLG", "PBHL", "PETR", "PHIL", "PHNU", "PHRM", "PHYL", "PHYS", "PPIA", "PSPA", "PSYC", "PSYT", "RCOD", "SART", "SHRP",
                "SOAN", "STAT", "THTR", "URDS", "URPL"};

            //  Courses is a list of list<string>, where each list<string> represents the data of a certain course.
            List<List<string>> Courses = new List<List<string>>();

            #region Selecting all the majors from the select list
            semester = new SelectElement(driver.FindElement(By.Name("p_term")));
            semester.SelectByValue(semesterNumber);
            driver.FindElement(By.TagName("form")).Submit();
            SelectElement majors = new SelectElement(driver.FindElement(By.Id("subj_id")));

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
            #endregion

            // extract the html source code of the page
            var html = driver.PageSource;

            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);

            // getting the title - crn - name - section of each course
            var htmlTitleNodes = htmlDoc.DocumentNode.SelectNodes("//th[@class='ddtitle']");

            // getting the whole div that contains the data of each course
            var htmlDataNodes = htmlDoc.DocumentNode.SelectNodes("//table[@class='datadisplaytable']/tbody");

            // getting started by specifying the order of the data to be represented.
            Courses.Add(new List<string> { "Title", "CRN", "Name", "Section", "Term", "Credits", "Actuall Enrollment", "Remaining Seats", "Linked Crns",
                "Time 1", "Days 1", "Place 1", "Schedule Type 1", "Instructors 1", "Time 2", "Days 2", "Place 2", "Schedule Type 2", "Instructors 2",
                "Time 3", "Days 3", "Place 3", "Schedule Type 3", "Instructors 3", "Time 4", "Days 4", "Place 4", "Schedule Type 4", "Instructors 4",
                "Time 5", "Days 5", "Place 5", "Schedule Type 5", "Instructors 5", "Time 6", "Days 6", "Place 6", "Schedule Type 6", "Instructors 6",
                "Time 7", "Days 7", "Place 7", "Schedule Type 7", "Instructors 7", "Time 8", "Days 8", "Place 8", "Schedule Type 8", "Instructors 8",
                "Time 9", "Days 9", "Place 9", "Schedule Type 9", "Instructors 9"});


            if (htmlDataNodes != null)
            {
                // splitting the html code according to the courses and storing them dataNodes
                var dataNodes = htmlDataNodes[0].InnerText.Split("\r\n\r\n\r\n\r\n\r\n\r\n\r\n\r\n");

                for (int i = 0; i < dataNodes.Length; i++)
                {
                    var item = dataNodes[i].Replace("Type", "").Replace("Time", "").Replace("Days", "").Replace("Where", "").Replace("Date Range", "").Replace("Schedule Type", "")
                        .Replace("Intructors", "").Replace("Scheduled Meeting Times", "").Replace("View Catalog Entry", "").Replace("Main Campus Campus", "").Replace("Course Syllabus", "");
                    var itemLines = item.Split("\r\n");

                    // courseData is the list where the information of the course are represented
                    List<string> courseData = new List<string>();
                    int start;
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
                    // the crn, title, name and section of the ith course have been added to courseData
                    for (int k = start; k < itemLines.Length; k++)
                    {
                        if (itemLines[k].Contains("Associated Term:"))
                            courseData.Add(itemLines[k].Replace("Associated Term: ", "")); // adding the semester
                        else if (itemLines[k].Contains("Credits"))
                        {
                            courseData.Add(itemLines[k].Replace("Credits", "")); // adding the credits
                            bool Found = false;
                            for (int w = 0; w < dynamicCourseDetails.Count; w++)
                            {
                                if (dynamicCourseDetails[w][0].ToString().Equals(courseData[1].ToString()))
                                {
                                    Found = true;
                                    courseData.Add(dynamicCourseDetails[w][1] != "" ? dynamicCourseDetails[w][1] : ""); // Actual Enrollment
                                    courseData.Add(dynamicCourseDetails[w][2] != "" ? dynamicCourseDetails[w][2] : ""); // Remaining Seats
                                    courseData.Add(dynamicCourseDetails[w][3] != "" ? dynamicCourseDetails[w][3] : ""); // Linked crns
                                    break;
                                }
                            }
                            if (!Found)
                            {
                                courseData.Add("");
                                courseData.Add("");
                                courseData.Add("");
                            }
                            // Actual Enrollment, Available Seats, and Linked Crns of the ith course have been added to courseData
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
                                    // Time, Days, Place, Schedule Type, Instructors of the ith course have been added to courseData
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine(itemLines[0]);
                                }
                            }
                            // Note: some courses may have up to 9 distinct timings (Time, Days, Place, Scedule Type, and Instructor), in such cases, all the timings will be saved next to each other.
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
                    // adding the data of the ith course to the list of courses
                }
            }

            #region Writing data to excel
            for (int i = 0; i < Courses.Count; i++)
            {
                for (int j = 0; j < Courses[i].Count; j++)
                {
                    string temp = Courses[i][j];
                    WriteToCell(i, j, Courses[i][j]);
                }
            }
            Save();
            #endregion

            driver.Close(); // closing the chrome driver
            watch.Stop(); // stopping the timer
            Console.WriteLine(watch.Elapsed); // Approx. Execution Time: 00:01:16.1272243
        }

        static string path = @"D:\Me\00. Classes\00. Programming\01. C#\04. Programs\Data Collector\Courses.xlsx";
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
            worksheet.Cells.AutoFitColumns(10, 100);
            package.Save();
        }
    }
}
