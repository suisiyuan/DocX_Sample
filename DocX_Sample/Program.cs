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
                Paragraph p = document.InsertParagraph("Here is Picture 1", false);
                Novacode.Image img = document.AddImage(@"../test.jpg");
                Picture pic = img.CreatePicture();
                p.InsertPicture(pic, 0);
                Console.WriteLine("pic.width: " + pic.Width);
                Console.WriteLine("pic.height: " + pic.Height);

                document.Save();
            }
            //阻塞
            Console.ReadLine();
        }
    }
}
