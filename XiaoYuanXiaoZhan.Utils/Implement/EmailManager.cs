using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using XiaoYuanXiaoZhan.Utils.Interface;
using XiaoYuanXiaoZhan.Utils.Model;

namespace XiaoYuanXiaoZhan.Utils.Implement
{
    public class EmailManager : IEmailManger
    {
        private static string _host;
        private static int _port;
        private static string _fromAddress;
        private static string _fromPassword;
        private static string _fromName;

        /// <summary>
        /// 设置默认的配置信息
        /// </summary>
        /// <param name="host"></param>
        /// <param name="port"></param>
        /// <param name="fromAddress"></param>
        /// <param name="fromPassword"></param>
        /// <param name="fromName"></param>
        public void SetDefaultConfig(string host, int port, string fromAddress, string fromPassword, string fromName)
        {
            _host = host;
            _port = port;
            _fromAddress = fromAddress;
            _fromPassword = fromPassword;
            _fromName = fromName;
        }

        /// <summary>
        /// 获取发送数据对象
        /// http://www.mimekit.net/docs/html/Introduction.htm
        /// </summary>
        /// <param name="emailModel"></param>
        /// <returns></returns>
        public MimeMessage GetMimeMessage(EmailModel emailModel,bool isUsingDefault = true)
        {
            if (isUsingDefault)
            {
                emailModel.MailFromAddress = _fromAddress;
                emailModel.MailFromUser = _fromName;
                emailModel.Host = _host;
                emailModel.Password = _fromPassword;
                emailModel.Port = _port;
            }

            MimeMessage message = new MimeMessage();
            message.From.Add(new MailboxAddress(emailModel.MailFromUser, emailModel.MailFromAddress));
            message.To.Add(new MailboxAddress(emailModel.MailToUser, emailModel.MailToAddress));
            message.Subject = emailModel.Subject;

            //邮件正文
            var alternativeBody = new MultipartAlternative();
            var textBody = new TextPart(TextFormat.Plain) { Text = emailModel.TextBody };
            var htmlBody = new TextPart(TextFormat.Html) { Text = emailModel.HtmlBody };
            alternativeBody.Add(textBody);
            alternativeBody.Add(htmlBody);
            Multipart multipart = new Multipart("mixed");
            multipart.Add(alternativeBody);

            //附件
            if (emailModel.Attachments != null && emailModel.Attachments.Count >= 0)
            {
                for (int i = 0; i < emailModel.Attachments.Count; i++)
                {
                    var path = emailModel.Attachments[i];
                    MimePart attachment = new MimePart()
                    {
                        Content = new MimeContent(File.OpenRead(path), ContentEncoding.Default),
                        ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                        ContentTransferEncoding = ContentEncoding.Base64,
                        FileName = Path.GetFileName(path)
                    };
                    multipart.Add(attachment);
                }
            }

            message.Body = multipart;
            return message;
        }

        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="emailModel"></param>
        /// <returns></returns>
        public async Task SendEmailAsync(EmailModel emailModel, bool isUsingDefault = true)
        {
            var message = GetMimeMessage(emailModel,isUsingDefault);
            using (SmtpClient client = new SmtpClient())
            {
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                client.Connect(emailModel.Host, emailModel.Port, SecureSocketOptions.Auto);
                client.Authenticate(emailModel.MailFromAddress, emailModel.Password);
                await client.SendAsync(message);
                client.Disconnect(true);
            }
        }

        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="emailModelList"></param>
        /// <param name="isUsingDefault"></param>
        /// <returns></returns>
        public async Task SendEmailAsync(List<EmailModel> emailModelList, bool isUsingDefault)
        {
            if (emailModelList.Count <= 0)
            {
                return;
            }
            using (SmtpClient client = new SmtpClient())
            {
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                for (int i = 0; i < emailModelList.Count; i++)
                {
                    var emailModel = emailModelList[i];
                    client.Connect(emailModel.Host, emailModel.Port, SecureSocketOptions.Auto);
                    client.Authenticate(emailModel.MailFromAddress, emailModel.Password);
                    await client.SendAsync(GetMimeMessage(emailModel));
                    client.Disconnect(true);
                }
            }
        }

        /// <summary>
        /// 批量发送邮件，要求发送者的账号、密码、服务器地址、端口是一致的
        /// </summary>
        /// <param name="emailModelList"></param>
        /// <param name="isUsingDefault"></param>
        /// <returns></returns>
        public async Task SendEmailBatch(List<EmailModel> emailModelList, bool isUsingDefault = true)
        {
            if (emailModelList.Count <= 0)
            {
                return;
            }

            var host = emailModelList[0].Host;
            var port = emailModelList[0].Port;
            var fromEmail = emailModelList[0].MailFromAddress;
            var pwd = emailModelList[0].Password;
            if (isUsingDefault)
            {
                host = _host;
                port = _port;
                fromEmail = _fromAddress;
                pwd = _fromPassword;
            }
            using (SmtpClient client = new SmtpClient())
            {
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                client.Connect(host, port, SecureSocketOptions.Auto);
                client.Authenticate(fromEmail, pwd);
                for (int i = 0; i < emailModelList.Count; i++)
                {
                    var emailModel = emailModelList[i];
                    await client.SendAsync(GetMimeMessage(emailModel));
                    Thread.Sleep(1500);
                }
                client.Disconnect(true);
            }
        }
    }
}

