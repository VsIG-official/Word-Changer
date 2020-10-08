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
		private const string ReadFrom = @"YourDisk:\Users\YourUser\Desktop\ReadFrom.txt";
        private const string ChangeIn = @"YourDisk:\Users\YourUser\Desktop\ChangeIn.docx";

        const int numberOfFonts = 4;
        const string FirstFont = "arial";
        const string SecondFont = "verdana";
        const string ThirdFont = "arial black";
        const string FourthFont = "times new roman";

		static readonly string[] Fonts = { FirstFont, SecondFont, ThirdFont, FourthFont};

        const int TextSize = 14;

        /// <summary>
        /// Defines the entry point of the application.
        /// </summary>
        static void Main()
        {
            Application app = new Application();
            Document doc = app.Documents.Add(Visible: true);
            Range rangeOfDoc = doc.Range();

            string phrase = File.ReadAllText(ReadFrom, Encoding.Default);
            rangeOfDoc.Text = phrase;

            int docLength = rangeOfDoc.Text.Length - 1;
            char[] letters = phrase.ToCharArray();
            rangeOfDoc.Font.Size = TextSize;

            Random random = new Random();

            for (int i = 0; i < docLength; i++)
			{
                Range tempR = doc.Range(i, docLength);
                int number = random.Next(numberOfFonts);

                tempR.Font.Name = Fonts[number];
            }

            doc.Save();
            app.Documents.Open(ChangeIn);
            Console.ReadKey();
        }
    }
}
