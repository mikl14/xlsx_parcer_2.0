using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace xlsx_parcer_2._0
{
    class Program
    {
        struct field
        {
            public string name;
            public string INN;
            public string[] type_nalog;
            public string[] nalog;

            public field(string Name, string inn, string[] Type_nalog, string[] Nalog)
            {
                name = Name;
                INN = inn;
                type_nalog = Type_nalog;
                nalog = Nalog;
            }
        }

        static void Main(string[] args)
        {


            List<field> fields = new List<field>();


            string paths = @".\files";
            string[] files = Directory.GetFiles(paths);

               Exel_field[] exel_fields = ReadExel("Export1.xlsx");

            Workbook wb = new Workbook("Export1.xlsx");

            // Получить все рабочие листы
            WorksheetCollection collection = wb.Worksheets;

            // Перебрать все рабочие листы
            Style style = new Style();

            int count_cells = 0;
                foreach (string file in files)
                {
                    string path = file;

                    for (int j = 0; j < Find_docs(path).Length; j++)
                    {
                      
                        for (int i = 1; i < Count_substr(Find_docs(path)[j], "НаимОрг") + 1; i++)
                        {
                            fields.Add(new field(Find_name(Find_docs(path)[j]),Find_INN(Find_docs(path)[j]), Find_type_nalog(Find_docs(path)[j]),Find_nalog(Find_docs(path)[j])));

                        }

                    }
               
                for(int i = 0; i < fields.ToArray().Length;i++)
                {
                    for (int j = 0; j < exel_fields.Length; j++)
                    {

                        if (fields.ToArray()[i].INN == exel_fields[j].INN)
                        {
                           
                           // Console.WriteLine(fields.ToArray()[i].name + " " + fields.ToArray()[i].INN );

                            Find_and_check_Exel(wb,collection,style, fields.ToArray()[i].INN, 0);


                        }
                    }
                      }
            }
            wb.Save("Export11.xlsx");

        }

        public static int Count_substr(String str, String substr)
        {
            int count = 0;
            int index = str.IndexOf(substr);
            while (index != -1)
            {
                count++;
                index = str.IndexOf(substr, index + 1);
            }
            return count;
        }
        public static int Count_Str(string str, char ch)
        {
            int count = 0;

            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] == ch)
                {
                    count++;
                }
            }

            return count;
        }

        public static string Find_name(string str)
        {
            try
            {
                str = str.Substring(str.IndexOf("НаимОрг=") + 8, str.IndexOf("ИНН") - 10);
                str = str.Replace("&quot;", "\"");
            }
            catch
            {

            }
            return str;

        }

        public static string Find_INN(string str)
        {

            try
            {

                str = str.Substring(str.IndexOf("ИННЮЛ=") + 7);
                str = str.Substring(0, str.IndexOf("\"/>"));
                str = str.Replace("&quot;", "\"");
            }
            catch
            {

            }
            return str;
        }

        public static string[] Find_nalog(string str)
        {

            string[] words = str.Split("СумУплНал=", StringSplitOptions.RemoveEmptyEntries);

            for (int i = 1; i < words.Length; i++)
            {
                try
                {
                    words[i] = words[i].Substring(1, words[i].IndexOf("\"/>") - 1);

                }
                catch
                {
                }

            }
            return words;
        }

        public static string[] Find_type_nalog(string str)
        {

            string[] words = str.Split("НаимНалог=\"", StringSplitOptions.RemoveEmptyEntries);

            for (int i = 1; i < words.Length; i++)
            {
                try
                {

                    words[i] = words[i].Substring(0, words[i].IndexOf("\""));

                    if (words[i] == "Налог, взимаемый в связи с  применением упрощенной  системы налогообложения")
                    {
                        words[i] = "УСНО";
                    }
                    else
                    {
                        words[i] = "";
                    }

                }
                catch
                {

                }
            }
            return words;
        }

        public static string[] Find_docs(string path)
        {

            using (StreamReader sr = new StreamReader(path))
            {
                string line;

                while ((line = sr.ReadLine()) != null)
                {

                    string[] words = line.Split("СведНП", StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 1; i < words.Length; i++)
                    {
                        try
                        {

                            words[i] = words[i].Substring(0, words[i].IndexOf("ДатаСост="));


                            // Console.WriteLine(words[i]);
                        }
                        catch
                        {
                            // Console.WriteLine("каво");
                        }

                    }
                    return words;

                }
            }


            return null;

            // Console.WriteLine("Hello World!");
        }

        public struct Exel_field
        {
            public string name;
            public string INN;
           

            public Exel_field(string Name, string inn)
            {
                name = Name;
                INN = inn;
            }
        }
         public static Exel_field[] ReadExel(string path)
        {
            List<Exel_field> exel_field = new List<Exel_field>();

            Workbook wb = new Workbook(path);

            // Получить все рабочие листы
            WorksheetCollection collection = wb.Worksheets;

            // Перебрать все рабочие листы
            for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
            {

                // Получить рабочий лист, используя его индекс
                Worksheet worksheet = collection[worksheetIndex];

                // Печать имени рабочего листа
                Console.WriteLine("Worksheet: " + worksheet.Name);

                int rows = worksheet.Cells.MaxDataRow;
                int cols = worksheet.Cells.MaxDataColumn;
                try
                {
                    for (int i = 0; i < rows; i++)
                    {

                        exel_field.Add(new Exel_field(worksheet.Cells[i, 2].Value.ToString(), worksheet.Cells[i, 0].Value.ToString()));

                    }
                    Console.WriteLine("Exel Ok");
                }
                catch (Exception e)
                {
                    Console.WriteLine("Ошибка чтения exel проверьте файл на наличие листов с странным форматированием");
                }
                

            }
            return exel_field.ToArray();
        }

        public static void Find_and_check_Exel(Workbook wb, WorksheetCollection collection, Style style, string name, int col)
        {
           

            for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
            {
            
                // Получить рабочий лист, используя его индекс
                Worksheet worksheet = collection[worksheetIndex];

                // Печать имени рабочего листа
                Console.WriteLine("Совпадение... " + name);
               
                // Получить количество строк и столбцов
                int rows = worksheet.Cells.MaxDataRow;
                for (int i = 0; i < rows; i++)
                {
                    // Console.WriteLine(worksheet.Cells[i, col].Value.ToString());
                        //Console.WriteLine(worksheet.Cells[i, col].Value.ToString());
                        if (name == worksheet.Cells[i, col].Value.ToString())
                        {
                            Cell cell = wb.Worksheets[worksheetIndex].Cells[i, col];

                            style = cell.GetStyle();
                            style.Font.Color = Color.Blue;
                            cell.SetStyle(style);
                            //Console.WriteLine("синий");
                            break;

                        }
                 
                   
                }

            }

      
        }
    }


}
