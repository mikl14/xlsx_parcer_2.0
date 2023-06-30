using System;
using System.Collections.Generic;
using System.IO;

namespace xlsx_parcer_2._0
{
    class Program
    {
        public struct Field // структура записи (имя, ИНН, массив всех типов налогов, суммы каждого налога)
        {
            public string name;
            public string INN;
            public string[] type_nalog;
            public string[] nalog;

               public Field(string Name, string inn, string[] Type_nalog, string[] Nalog)   
               {
                   name = Name;
                   INN = inn;
                   type_nalog = Type_nalog;
                   nalog = Nalog;
               }

        }


        public static StreamWriter sw = new StreamWriter("Test.txt"); // стрим райтер закрывать только когда он заполнен всем нужным

        static void Main(string[] args)
        {


            List<Field> fields = new List<Field>();  //список всех записей


            string paths = @".\files"; //путь к xml
            string[] files = Directory.GetFiles(paths);

        
           
                foreach (string file in files) //перебераем все файлы
                {

               
                    string path = file;
                    fields.Clear();
                    String[] Find_doc = Find_docs(path); //парсим все отдельные записи (массив сырых записей)

                    for (int j = 0; j < Find_doc.Length; j++)
                    {
                      
                        for (int i = 1; i < Count_substr(Find_doc[j], "НаимОрг") + 1; i++) // бля я слишком ватный не могу вспомнить почему такая проверка но она важная епт
                        {
                           fields.Add(new Field(Find_name(Find_doc[j]),Find_INN(Find_doc[j]),Find_type_nalog(Find_doc[j]),Find_nalog(Find_doc[j]))); // заполняем список
                        
                        }

                    }

                    foreach (Field pole in fields) // перебор списка
                    {
                        show_stats(pole); // вывод в консоль всей записи списка
                        add_to_text(pole); // вывод в текстовик всей записи списка
                    }

                }
            sw.Close(); // закрыть врайтер файлика
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
        } // считает число подстрок в строке
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
        } // считает число чаров в строке

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

        } // ищет имя в сырой записи

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
        } // ищет инн в сырой записи

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
        } // ищет суммы в сырой записи

        public static string[] Find_type_nalog(string str)
        {

            string[] words = str.Split("НаимНалог=\"", StringSplitOptions.RemoveEmptyEntries);

            for (int i = 1; i < words.Length; i++)
            {
                try
                {
                    words[i] = words[i].Substring(0, words[i].IndexOf("\""));
            
                }
                catch
                {

                }
            }
            return words;
        } // ищет типы налогов в сырой записи

        public static string[] Find_docs(string path) // ищет делает массив сырых записей
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

                            words[i] = words[i].Substring(0, words[i].IndexOf("ДатаСост=")); //нет он не пропускает последнюю запись, хотя козалось бы... ну или я стал очень глупый чтобы понять всю соль 

                        }
                        catch
                        {
                        }

                    }
                    return words;

                }
            }


            return null;

            // Console.WriteLine("Hello World!");
        }

        


        public static void show_stats(Field field) // вывод в консоль всей записи списка
        {
            for(int i = 1; i < field.type_nalog.Length;i++)
            {
                Console.WriteLine(field.INN + " " + field.name +" " + field.nalog[i] + " " + field.type_nalog[i]);

            }
        } 

       
        public static void add_to_text(Field field) // вывод в текстовик всей записи списка
        {
            try
            {

               
                for (int i = 1; i < field.type_nalog.Length; i++)
                {
                    sw.WriteLine(field.INN + " " + field.name + " " + field.nalog[i] + " " + field.type_nalog[i]);
                }
                
            }
            catch (Exception e)
            {
                Console.WriteLine("Panic !!!!!");
            }


        }
    }


}
