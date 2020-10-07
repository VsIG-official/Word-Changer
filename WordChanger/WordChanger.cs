using System;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace WordChanger
{
	/// <summary>
	/// Main class (hope its singletone)
	/// </summary>
	class WordChanger
    {
		private const string DestinationToYourFile = @"C:\Users\Valentin\Desktop\text";
		private const string txt = ".txt";
		private const string docx = ".docx";

		/// <summary>
		/// Defines the entry point of the application.
		/// </summary>
		static void Main()
        {
            Application app = new Application();
            Document doc = app.Documents.Add(Visible: true);
            Range r = doc.Range();
            string phrase = File.ReadAllText(@"DestinationToYourFile + txt", Encoding.Default);
            string firstFont = "arial";
            string secondFont = "verdana";
            string thirdFont = "arial black";
            string fourthFont = "times new roman";
            r.Text = phrase;
            int docLength = r.Text.Length - 1;
            char[] letters = phrase.ToCharArray();
            r.Font.Size = 14;
            Random random = new Random();

            for (int i = 0; i < docLength; i++)
			{
                Range tempR = doc.Range(i, docLength);
                int number = random.Next(4);
                if (number == 1)
                {
                    tempR.Font.Name = firstFont;
                }
                else if(number == 2)
                {
                    tempR.Font.Name = secondFont;
                }
                else if (number == 3)
                {
                    tempR.Font.Name = thirdFont;
                }
				else
				{
                    tempR.Font.Name = fourthFont;
                }
            }

            doc.Save();
            app.Documents.Open(@"DestinationToYourFile + docx");
            Console.ReadKey();
        }
    }
}
