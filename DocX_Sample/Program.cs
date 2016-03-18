using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;

namespace DocX_Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            //如果已经有文档则删除
            if (File.Exists("../test.docx"))
            {
                File.Delete("../test.docx");
            }

            using (DocX document = DocX.Create(@"../test.docx"))
            {   
                //添加文字
                Formatting formatting = new Formatting();
                formatting.Bold = true;
                formatting.FontColor = Color.Red;
                formatting.Size = 30;
                document.InsertParagraph("test!", false, formatting);

                //添加图片
                Paragraph p = document.InsertParagraph();
                Novacode.Image img = document.AddImage(@"../test.jpg");
                Picture pic = img.CreatePicture();
                p.InsertPicture(pic, 0);
                Console.WriteLine("照片宽度：" + pic.Width);
                Console.WriteLine("照片高度： " + pic.Height);


                //添加表格
                Table table = p.InsertTableAfterSelf(3, 3);


                //页眉页脚控制
                document.AddFooters();
                Footers footers = document.Footers;
                Footer first = footers.first;
                document.DifferentFirstPage = true;
                document.DifferentOddAndEvenPages = true;
                p = first.InsertParagraph();
                p.Append("This is the first pages footer.");


                //输出文档属性
                Console.WriteLine("段落数：" + document.Paragraphs.Count);
                Console.WriteLine("图片数：" + document.Pictures.Count);
                Console.WriteLine("节数：" + document.Sections.Count);

                document.Save();
            }
            //阻塞
            Console.ReadLine();
        }
    }
}
