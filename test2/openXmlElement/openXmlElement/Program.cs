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
            try
            {
                using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(fileName, true))
                {
                    // 获取主体
                    Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

                    // 获取所有元素
                    IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> elements = body.Elements<DocumentFormat.OpenXml.OpenXmlElement>();
                    //遍历元素列表，输出内容
                    foreach (DocumentFormat.OpenXml.OpenXmlElement element in elements)
                    {
                        Console.WriteLine(element.InnerText);
                    }
                    Console.ReadKey();
                }
            }
            catch
            {
                Console.WriteLine("没有找到文件");
                Console.ReadKey();
            }
            
        }
    }
}
