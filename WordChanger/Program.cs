using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
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
            r.Text = "Hello world";
            int docLength = r.Text.Length - 1;
            r.Text = docLength.ToString();
            Random random = new Random();
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

			for (int i = 0; i < docLength; i++)
			{
                Console.WriteLine("Line");
			}
            Console.ReadKey();
        }
    }
}
