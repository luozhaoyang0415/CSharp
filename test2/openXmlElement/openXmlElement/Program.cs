using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace openXmlElement
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadWord readWord = new ReadWord(@"F:\工作\github\CSharp\test2\科研细则.docx");
            foreach(string i in readWord.TextList)
            {
                Console.WriteLine(i);
            }
            
            Console.ReadKey();

        }
    }
}
