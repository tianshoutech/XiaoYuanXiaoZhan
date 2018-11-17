using System;
using System.Collections.Generic;
using System.Text;

namespace XiaoYuanXiaoZhan.Utils.Model
{
    public class EmailModel
    {
        public string MailFromUser { get; set; }
        public string MailFromAddress { get; set; }
        public string Host { get; set; }
        public int Port { get; set; }
        public string Password { get; set; }

        public string MailToUser { get; set; }
        public string MailToAddress { get; set; }
        public string CC { get; set; }
        public string BCC  { get; set; }
        public string Subject { get; set; }
        public List<string> Attachments { get; set; }
        public string HtmlBody { get; set; }
        public string TextBody { get; set; }
        public int SendServerType { get; set; }
    }
}
