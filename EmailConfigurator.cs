using Newtonsoft.Json;
using SendGrid.Helpers.Mail;
using SendGrid;
using System.Data;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace EmailSender
{
    public class EmailConfigurator
    {
        private EmailSettings Settings { get; set; }
        private static int SendEmailCount { get; set; }
        private static int FailEmailCount { get; set; }

        public EmailConfigurator()
        {
            SendEmailCount = 0;
            FailEmailCount = 0;
            Settings = JsonConvert.DeserializeObject<EmailSettings>(
                File.ReadAllText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "settings.json")));
        }

        public async Task SendEmailsToUsers()
        {
            var dataSet = datasetImportFromExcel(Settings.EmailExcelPath, Settings.ExcelHeader);

            foreach (DataRow row in dataSet.Tables[Settings.ExcelSheetName].Rows)
            {
                object[] rowData = row.ItemArray;
                if (rowData[2] == DBNull.Value)
                    break;
                await sendEmail((string)rowData[2], (string)rowData[1], (string)rowData[6], (string)rowData[7]);
            }
            Console.Out.WriteLine("SendEmailCount: " + SendEmailCount);
            Console.Out.WriteLine("FailEmailCount: " + FailEmailCount);
            SendEmailCount = 0;
            FailEmailCount = 0;
        }

        private async Task sendEmail(string sendToEmailAddress, string sendToName, string emailSubject, string emailText)
        {
            var client = new SendGridClient(Settings.SmtpPassword);
            var msg = new SendGridMessage()
            {
                From = new EmailAddress(Settings.FromEmail, Settings.FromName),
                Subject = emailSubject,
                PlainTextContent = emailText,
            };
            msg.AddTo(new EmailAddress(sendToEmailAddress, sendToName));
            var response = await client.SendEmailAsync(msg);
            if(response.IsSuccessStatusCode)
            {
                SendEmailCount++;
                Console.Out.WriteLine("Email queued successfully! Email Address: " + sendToEmailAddress);
            } else
            {
                FailEmailCount++;
                Console.Out.WriteLine("Something went wrong! Email Address: " + sendToEmailAddress);
            }
        }

        private DataSet datasetImportFromExcel(string FilePath, bool headers)
        {
            var _xl = new Excel.Application();
            var wb = _xl.Workbooks.Open(FilePath);
            var sheets = wb.Sheets;
            DataSet? dataSet = null;
            if (sheets != null && sheets.Count != 0)
            {
                dataSet = new DataSet();
                foreach (var item in sheets)
                {
                    var sheet = (Excel.Worksheet)item;
                    DataTable? dt = null;
                    if (sheet != null)
                    {
                        dt = new DataTable();
                        dt.TableName = sheet.Name;
                        var ColumnCount = ((Excel.Range)sheet.UsedRange.Rows[1, Type.Missing]).Columns.Count;
                        var rowCount = ((Excel.Range)sheet.UsedRange.Columns[1, Type.Missing]).Rows.Count;

                        for (int j = 0; j < ColumnCount; j++)
                        {
                            var cell = (Excel.Range)sheet.Cells[1, j + 1];
                            var column = new DataColumn(headers ? (string)cell.Value : string.Empty);
                            dt.Columns.Add(column);
                        }

                        for (int i = 0; i < rowCount; i++)
                        {
                            var r = dt.NewRow();
                            for (int j = 0; j < ColumnCount; j++)
                            {
                                var cell = (Excel.Range)sheet.Cells[i + 1 + (headers ? 1 : 0), j + 1];
                                r[j] = cell.Value;
                            }
                            dt.Rows.Add(r);
                        }

                    }
                    dataSet.Tables.Add(dt);
                }
            }
            _xl.Quit();
            return dataSet;
        }
    }
}
