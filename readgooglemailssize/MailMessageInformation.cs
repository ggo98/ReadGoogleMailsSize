using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace readgooglemailssize
{
    internal class MailMessageInformation
    {
        public string Subject { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public DateTime Date { get; set; }
        public Int64 Size { get; set; }
        public string Snippet { get; set; }
        public string Labels { get; set; }
    }
}
