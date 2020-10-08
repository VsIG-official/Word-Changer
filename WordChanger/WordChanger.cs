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
		private const string readFrom = @"C:\Users\Valentin\Desktop\readFrom.txt";
        private const string changeIn = @"C:\Users\Valentin\Desktop\changeIn.docx";

        static string FirstFont = "arial";
        static string SecondFont = "verdana";
        static string ThirdFont = "arial black";
        static string FourthFont = "times new roman";

		static string[] Fonts = { FirstFont, SecondFont, ThirdFont, FourthFont};

        const int TextSize = 14;

        /// <summary>
        /// Defines the entry point of the application.
        /// </summary>
        static void Main()
        {
            Application app = new Application();
            Document doc = app.Documents.Add(Visible: true);
            Range rangeOfDoc = doc.Range();

            string phrase = File.ReadAllText(readFrom, Encoding.Default);
            rangeOfDoc.Text = phrase;

            int docLength = rangeOfDoc.Text.Length - 1;
            char[] letters = phrase.ToCharArray();
            rangeOfDoc.Font.Size = TextSize;

            Random random = new Random();

            for (int i = 0; i < docLength; i++)
			{
                Range tempR = doc.Range(i, docLength);
                int number = random.Next(4);

                tempR.Font.Name = Fonts[number];
            }

            doc.Save();
            app.Documents.Open(changeIn);
            Console.ReadKey();
        }
    }
}
