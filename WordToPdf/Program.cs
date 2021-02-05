using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;

using System.ComponentModel;

using System.Drawing;
namespace WordToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            
            DateTime beforDT = System.DateTime.Now; 
            for (int i = 0; i < args.Length; i++)
            {
                Console.WriteLine(args[i]);
            }
         

            //测试环境注释掉下方命令行接收代码
            if (args.Length<3)
            {
                Console.WriteLine("错误操作!");
                Console.WriteLine("-type    word      excel");
                Console.WriteLine("fromPath");
                Console.WriteLine("toPath");
                return;
            }

            String type = args[0];
            String from = args[1];
            String to = args[2];

            //String type = "-execl";
            //String from = @"C:\Users\Administrator\Desktop\zcfzb.xls";
            //String to = @" C:\Users\Administrator\Desktop\test.pdf";


            Console.WriteLine("正在跑代码!");
    
            switch (type)
            {
                 //文件后缀名.docx  
                case "-word" :
                    new Wlword().WordToPDF(from, to);
                    break;

                //文件后缀名为xls
                case "-execl":
                    new WlExecl().ExeclToPDF(from, to);
                    break;

                //文件后缀名为ppt
                case "-ppt":
                    new Wlppt().pptToPdf(from, to);
                    break;

                default :
                    Console.WriteLine("输入有误!");
                    break;
            }

            DateTime afterDT = System.DateTime.Now;
            TimeSpan ts = afterDT.Subtract(beforDT);
            Console.WriteLine("DateTime总共花费{0}ms.", ts.TotalMilliseconds); 
            
        }

    }
}
