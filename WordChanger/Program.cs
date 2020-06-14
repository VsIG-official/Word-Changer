using System;
using Microsoft.Office.Interop.Word;

namespace WordChanger
{
	class Program
	{
        static void Main(string[] args)
        {
            Application app = new Application();
            Document doc = app.Documents.Add(Visible: true);
            Range r = doc.Range();
            string phrase = "Hello world";
            string firstFont = "arial";
            string secondFont = "verdana";
            r.Text = phrase;
            int docLength = r.Text.Length - 1;
            r.Text = docLength.ToString();
            char[] letters = phrase.ToCharArray();
            r.Text = letters.ToString();
            r.Font.Size = 14;
            Random random = new Random();

			for (int i = 0; i < docLength; i++)
			{
                Console.WriteLine(letters[i]);
            }

            foreach (var letter in letters)
            {
                int number = random.Next(2);
                if (number == 1)
                {
                    r.Font.Name = firstFont;
                }
                else
                {
                    r.Font.Name = secondFont;
                }
            }

            doc.Save();
            app.Documents.Open(@"C:\Users\Valentin\Desktop\ForFormatting.docx");
            Console.ReadKey();

			try
			{
                doc.Close();
                app.Quit();
			}
			catch (Exception e)
			{
                Console.WriteLine(e.Message);
			}


            Console.ReadKey();
        }
    }
}
