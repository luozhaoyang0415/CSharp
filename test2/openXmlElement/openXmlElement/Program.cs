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
            string fileName = @"F:\工作\github\CSharp\test2\科研细则.docx";
            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(fileName, false))
            {
                // 获取主体
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

                // 获取所有段落
                IEnumerable<Paragraph> paragraphs =body.Elements<Paragraph>();
                //遍历段落，输出内容
                foreach (Paragraph paragraph in paragraphs)
                {
                    Console.WriteLine(paragraph.InnerText);
                }
                Console.ReadKey();
            }
        }
    }
}
