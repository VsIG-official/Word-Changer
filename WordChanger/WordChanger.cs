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
        // Path to file, from which You want to read
		private const string ReadFrom = @"YourDisk:\Users\YourUser\Desktop\ReadFrom.txt";
        // Path to file, where You want to save Your text with different fonts
        private const string ChangeIn = @"YourDisk:\Users\YourUser\Desktop\ChangeIn.docx";

        // All possible fonts and their count
        const int numberOfFonts = 4;
        const string FirstFont = "arial";
        const string SecondFont = "verdana";
        const string ThirdFont = "arial black";
        const string FourthFont = "times new roman";

		static readonly string[] Fonts = { FirstFont, SecondFont, ThirdFont, FourthFont};

        // Size of a text
        const int TextSize = 14;

        /// <summary>
        /// Defines the entry point of the application.
        /// </summary>
        static void Main()
        {
            // Create range of .docx file
            Application app = new Application();
            Document doc = app.Documents.Add(Visible: true);
            Range rangeOfDoc = doc.Range();

            // Get the text from .txt file and set text of doc to it
            string phrase = File.ReadAllText(ReadFrom, Encoding.Default);
            rangeOfDoc.Text = phrase;

            // Count length of text (in characters)
            int docLength = rangeOfDoc.Text.Length - 1;
            // Put all chars to an array
            char[] letters = phrase.ToCharArray();
            // Set text to a specific size
            rangeOfDoc.Font.Size = TextSize;

            // Create new random
            Random random = new Random();

            // Iterate through each character
            for (int i = 0; i < docLength; i++)
			{
                Range tempR = doc.Range(i, docLength);
                // Generate random number
                int number = random.Next(numberOfFonts);

                // Set the random font to letter
                tempR.Font.Name = Fonts[number];
            }

            // Save the .docx and open it
            doc.Save();
            app.Documents.Open(ChangeIn);
            Console.ReadKey();
        }
    }
}
