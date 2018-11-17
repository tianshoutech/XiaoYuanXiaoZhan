using System;
using System.Collections.Generic;
using XiaoYuanXiaoZhan.Utils.Implement;
using XiaoYuanXiaoZhan.Utils.Model;

namespace Tests
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExcelTest.GenerateExcel();

            var emailManger = new EmailManager();
            var emailModel = new EmailModel();
            emailModel.Subject = "测试";
            emailModel.HtmlBody = "<p>测试</p>";
            emailModel.TextBody = "测试";
            emailModel.MailToAddress = "gebizhuifengren@aliyun.com";
            emailModel.MailToUser = "戈壁追风人";
            emailModel.Attachments = new List<string>();
            emailModel.Attachments.Add(@"C:\Users\stron\Desktop\ddd.txt");
            emailModel.Attachments.Add(@"C:\Users\stron\Desktop\bbb.xlsx");

            var list = new List<EmailModel>();
            for (int i = 0; i < 30; i++)
            {
                list.Add(emailModel);
            }

            emailManger.SendEmailBatch(list).Wait();

            Console.WriteLine("生成完毕");
            Console.ReadKey();
        }
    }
}
