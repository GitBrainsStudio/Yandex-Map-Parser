using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace YaMapParse
{
    class Program
    {
        public static string appRoot;

        static void UnhandeldExceptionTrapper(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                ErrorLog(e.ExceptionObject.ToString());
            }
            catch
            {

            }
            finally
            {
                try
                {
                    CloseApp();
                }

                catch
                {

                }

            }


        }

        static void Main(string[] args)
        {

            try
            {
                EventLog("***********************************************");
                EventLog("Запуск программы");
                System.AppDomain.CurrentDomain.UnhandledException += UnhandeldExceptionTrapper;

                string btn_search = ConfigurationManager.AppSettings["btn_search"];
                string input = ConfigurationManager.AppSettings["input"];
                string title = ConfigurationManager.AppSettings["title"];
                string descr = ConfigurationManager.AppSettings["descr"];
                string capcha = ConfigurationManager.AppSettings["capcha"];
                string capcha_text = ConfigurationManager.AppSettings["capcha_text"];

                string url = ConfigurationManager.AppSettings["url"];

                appRoot = AppDomain.CurrentDomain.BaseDirectory;

                ChromeDriver chrome = new ChromeDriver();

                var list = Get();
                int count = list.Count();
                if (count == 0) throw new Exception("В документе нет данных для проверки");
                EventLog("Прочитаны строки из файла. Общее количество: " + count);
                int index = 1;

                chrome.Navigate().GoToUrl(url);

                EventLog("Переход на ссылку");
                foreach (var i in list)
                {
                    try
                    {
                        chrome.FindElementByXPath(input).SendKeys(i.ToString());
                        chrome.FindElementByXPath(btn_search).Click();

                        while (true)
                        {
                            try
                            {
                                chrome.FindElementByXPath(btn_search);
                                break;
                            }
                            catch
                            {
                                try
                                {
                                    if (chrome.FindElementByXPath(capcha).Text == capcha_text)
                                    {
                                        EventLog("Обнаружена капча");
                                        Console.WriteLine("Обнаружена капча. Введите её вручную и нажмите любую клавишу в консоле для продолжения цикла.");
                                        Console.ReadKey();
                                        throw new ArgumentNullException("Обнаружена капча");
                                    }
                                }

                                catch (ArgumentNullException e)
                                {
                                    throw new Exception(e.Message);
                                }

                                catch
                                {

                                }
                            }
                        }

                        string _title;
                        string _descr;

                        while (true)
                        {
                            try
                            {
                                _title = chrome.FindElementByXPath(title).Text;
                                break;
                            }
                            catch
                            {

                            }
                        }

                        while (true)
                        {
                            try
                            {
                                _descr = chrome.FindElementByXPath(descr).Text;
                                break;
                            }
                            catch
                            {

                            }
                        }


                        StatusLine(i.ToString() + " : " + index + " : 1");
                        Write(i.ToString() + ";" + _title + " " + _descr);
                        Console.WriteLine("Обработана строка " + index + " из " + count + " : успешно");
                    }
                    catch (Exception ex)
                    {
                        Write(i.ToString() + ";" + "Ошибка при получении данных");
                        StatusLine(i.ToString() + " : " + index + " : 0");
                        ErrorLog(ex.Message);
                        Console.WriteLine("Обработана строка " + index + " из " + count + " : ошибка");
                    }

                    finally
                    {
                        index++;
                        chrome.FindElementByXPath(input).SendKeys(Keys.Control + "a");
                        chrome.FindElementByXPath(input).SendKeys(Keys.Delete);
                    }
                }
                EventLog("Программа отработала успешно");
            }

            catch (Exception ex)
            {
                ErrorLog(ex.Message);
                EventLog("Программа отработана с ошибкой");
            }

            finally
            {
                CloseApp();
            }
        }

        public static List<object> Get()
        {
            var list = new List<object>();
            using (StreamReader streamReader = new StreamReader(appRoot + "координаты на проверку.txt", encoding: Encoding.GetEncoding(1251)))
            {
                string line;
                while (!streamReader.EndOfStream)
                {
                    line = streamReader.ReadLine();
                    list.Add(line);
                }
            }

            return list;
        }

        public static void Write(string t)
        {
            using (StreamWriter streamWriter = new StreamWriter(appRoot + "результат проверки.txt", append: true, encoding: Encoding.GetEncoding(1251)))
            {
                streamWriter.WriteLine(t);
            }
        }

        public static void EventLog(string eventLog)
        {
            using (StreamWriter streamWriter = new StreamWriter(appRoot + "события.txt", append: true, encoding: Encoding.GetEncoding(1251)))
            {
                streamWriter.WriteLine(DateTime.Now + " : " + eventLog);
            }
        }
        public static void StatusLine(string statusLine)
        {
            using (StreamWriter streamWriter = new StreamWriter(appRoot + "статуслайны.txt", append: true, encoding: Encoding.GetEncoding(1251)))
            {
                streamWriter.WriteLine(statusLine);
            }
        }
        public static void ErrorLog(string Exception)
        {
            using (StreamWriter streamWriter = new StreamWriter(appRoot + "ошибки.txt", append: true, encoding: Encoding.GetEncoding(1251)))
            {
                streamWriter.WriteLine(DateTime.Now + " : " + Exception);
                streamWriter.WriteLine();
            }
        }

        public static void KillProccess(string proccessName)
        {
            Process[] processList = Process.GetProcessesByName(proccessName);

            foreach (var process in processList)
            {
                process.Kill();
            }
        }

        public static void CloseApp()
        {
            try
            {
                KillProccess("chrome");
                KillProccess("chromedriver");
            }
            catch
            {

            }
            finally
            {
                Environment.Exit(1);
            }
        }

    }
    }

