using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;
using System.IO;

namespace ExzelHandler
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            if(args.Length == 2)
            {
                Console.Write("args.Length == 2 \n");

                Console.Write("Длина = " + args.Length + "\n");

                Console.Write("Press any key to continue \n");
                Console.ReadKey();

                Console.Write(args[0] + "\n");

                Console.Write("Press any key to continue \n");
                Console.ReadKey();

                Console.Write(args[1] + "\n");
            }
            else
            {
                Console.Write("args.Length != 2 \n");

                Console.Write(args.Length + "\n");
            }

            Console.Write("Press any key to continue \n");
            Console.ReadKey();

            List<Person> persons = new List<Person>();

            string filePath = args[0];

            Console.Write("Путь преобразован \n");
            Console.ReadKey();

            persons = JsonConvert.DeserializeObject<List<Person>>(args[1]);

            Console.Write("Список преобразован \n");
            Console.ReadKey();


            ReadExcelFile(filePath, persons);

            Console.ReadKey();
        }

        public static void ReadExcelFile(string filePath, List<Person> persons)
        {
            // Открытие существующего файла Excel
            if (File.Exists(filePath))
            {
                try
                {
                    var workbook = new XLWorkbook(filePath);

                    // Продолжайте работу с workbook

                    Console.Write("Экзель файлзапушен \n");
                    Console.ReadKey();

                    // Получение листа по имени
                    var worksheet = workbook.Worksheet("Данные");

                    Console.Write("Лист данные получен \n");
                    Console.ReadKey();

                    for (int i = 0; i < persons.Count; i++)
                    {
                        var bumber = i + 1;

                        worksheet.Cell("A" + bumber).Value = persons[i].name;
                        worksheet.Cell("B" + bumber).Value = persons[i].age;
                        worksheet.Cell("C" + bumber).Value = persons[i].money;

                        Console.Write(persons[i].name + " " + persons[i].age + " " + persons[i].money);
                    }

                    // Сохранение изменений в файл
                    workbook.Save();

                    Console.Write("Проверьте экзель файл на наличие изменений \n");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Произошла ошибка при открытии файла: " + ex.Message);
                }
            }
            else
            {
                Console.WriteLine("Файл не найден: " + filePath);
            }
        }

        public class Person
        {
            public string name;
            public int age;
            public int money;
        }
    }
}
