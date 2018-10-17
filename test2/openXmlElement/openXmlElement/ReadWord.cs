using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace openXmlElement
{
    class ReadWord
    {
        public string Text = "";
        public ReadWord(string filename)
        {
            try
            {
                using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(filename, true))
                {
                    // 获取主体
                    DocumentFormat.OpenXml.Wordprocessing.Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

                    // 获取所有元素
                    IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> elements = body.Elements<DocumentFormat.OpenXml.OpenXmlElement>();
                    //遍历元素列表，输出内容
                    foreach (DocumentFormat.OpenXml.OpenXmlElement element in elements)
                    {
                        Text += element.InnerText+ "\r\n";
                        //Console.WriteLine(element.InnerText);
                    }
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
