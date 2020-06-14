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
            r.Text = phrase;
            r.Font.Size = 14;
            Random random = new Random();



            for (int i = 1; i < docLength; i++)
			{
				for (int j = 2; j < docLength; j++)
				{
                    Range tempR = doc.Range(0,4);

                    int number = random.Next(2);

                    if (number == 1)
                    {
                        tempR.Font.Name = firstFont;
                    }
                    else
                    {
                        tempR.Font.Name = secondFont;
                    }
                }

                //string text = doc.Words[i].Text;
                //Console.WriteLine("Word {0} = {1}", i, text);

                //Range rng = doc.Content;
                //rng.Select();
                //Console.WriteLine("Characters: " + doc.Characters.Count.ToString());

                //Range tempR=
                //letters[i]

                //letters[i];
            }

            foreach (var letter in letters)
            {


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
