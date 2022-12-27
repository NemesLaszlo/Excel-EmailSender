namespace EmailSender
{
    public class EmailSettings
    {
        public string SmtpHost { get; set; }
        public int SmtpPort { get; set; }
        public bool SslEnabled { get; set; }
        public string SmtpUser { get; set; }
        public string SmtpPassword { get; set; }
        public string SmtpDomain { get; set; }
        public int Timeout { get; set; }
        public string FromEmail { get; set; }
        public string FromName { get; set; }
        public string EmailExcelPath { get; set; }
        public string ExcelSheetName { get; set; }
        public bool ExcelHeader { get; set; }
    }
}
