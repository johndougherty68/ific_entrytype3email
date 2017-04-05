using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Net;
using System.Net.Mail;
using System.Net.Mime;
//using System.Threading;
using System.ComponentModel;
using NLog;
using System.Xml.Linq;
using System.IO;

namespace EntryType3Email
{
    class Program
    {
        private static string smtpServer;
        private static string toEmails;
        private static string fromEmail;
        private static string creds;
        private static string singleBondToCheck = "";
        private static string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        private static bool simulateOnly = false;
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private static int numDays = 30;
        static string body;
        static void Main(string[] args)
        {
            try
            {

                config();
                DateTime dateFrom = DateTime.Now.AddDays(numDays * -1);
                cbpmqdbEntities1 db = new cbpmqdbEntities1();
                //Find AS records with an entry type of 3 

                var q = from p in db.vType3EntriesWithBO where p.file_date >= dateFrom select p;

                if (singleBondToCheck != "")
                {
                    q = (from p in db.vType3EntriesWithBO
                         where p.file_date >= dateFrom && p.entry_type == 3
                         && p.bond_number.Contains(singleBondToCheck)
                         select p).Take(1);
                }

                logger.Info("Found " + q.Count() + " AS Records");
                //For each record
                foreach (vType3EntriesWithBO asr in q.ToList())
                {
                    logger.Info("Processing Bond Number " + asr.bond_number +
                          ", Entry Number " + asr.entry_number1 +
                          ", Entry Date " + asr.entry_date + ", Effective Date " +
                          asr.bond_effective_date + ", Importer " +
                          asr.importer_of_record + ", File date " + asr.file_date.ToShortDateString());

                    //See if there's already an email sent for it
                    var e = from p in db.ASAQEmails
                            where
                                p.BondNumber == asr.bond_number
                                && p.EntryDate == asr.entry_date
                                && p.EffectiveDate == asr.bond_effective_date
                                && p.Importer == asr.importer_of_record
                                && p.EntryNumber == asr.entry_number1
                                && (p.FileDate == asr.file_date || p.FileDate == null)
                            select p;
                    if (e.Count() > 0)
                    {
                        logger.Info("Email for Bond Number " + asr.bond_number +
                          ", Entry Date " + asr.entry_date + ", Effective Date " +
                            ", Entry Number " + asr.entry_number1 +
                          asr.bond_effective_date + ", Importer " +
                          asr.importer_of_record + ", File date " + asr.file_date.ToShortDateString() + " already sent.");
                        continue;
                    }

                    //Create and send email
                    logger.Info("Email for Bond Number " + asr.bond_number +
                        ", Entry Date " + asr.entry_date + ", Effective Date " +
                          ", Entry Number " + asr.entry_number1 +
                        asr.bond_effective_date + ", Importer " +
                        asr.importer_of_record + ", File date " + asr.file_date.ToShortDateString() + " has NOT BEEN sent.");
                    //for now, we'll just write to the console
                    logger.Info(asr.importer_of_record + " " + asr.entry_date.ToString());
                    //logger.Info("Sendmail has been commented out!");

                    sendMail(asr, simulateOnly);

                    if (!simulateOnly)
                    {

                        //Add "mail sent" record for this AS record
                        ASAQEmail asaq = new ASAQEmail();
                        asaq.BondNumber = asr.bond_number;
                        asaq.EntryDate = asr.entry_date;
                        asaq.EffectiveDate = asr.bond_effective_date;
                        asaq.Importer = asr.importer_of_record;
                        asaq.EmailSentDate = DateTime.Now;
                        asaq.EntryNumber = asr.entry_number1;
                        asaq.FileDate = asr.file_date;
                        db.ASAQEmails.Add(asaq);
                        db.SaveChanges();
                    }
                }

            }
            catch (Exception e)
            {
                logger.Error(e.Message);
                return;
            }
        }

        private static void config()
        {
            string appConfigLoc = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;
            if (!File.Exists(appConfigLoc))
            {
                throw new FileNotFoundException("Config file " + appConfigLoc + " not found. Exiting.");
            }

            XDocument doc = XDocument.Load(appConfigLoc);

            // Get the appSettings element as a parent
            XContainer appSettings = doc.Element("configuration").Element("appSettings");

            // step through the "add" elements
            foreach (XElement xe in appSettings.Elements("add"))
            {
                // Get the values
                string addKey = xe.Attribute("key").Value;
                string addValue = xe.Attribute("value").Value;

                switch (addKey.ToLower())
                {
                    case "singlebondtocheck":
                        logger.Info("Setting singleBondToCheck to " + addValue);
                        singleBondToCheck = addValue;
                        break;
                    case "smtpserver":
                        logger.Info("Setting SMTP server to " + addValue);
                        smtpServer = addValue;
                        break;
                    case "simulateonly":
                        logger.Info("Setting simulateOnly to " + addValue);
                        simulateOnly = bool.Parse(addValue);
                        break;
                    case "numdays":
                        logger.Info("Setting numDays to " + addValue);
                        numDays = int.Parse(addValue);
                        break;
                    case "creds":
                        logger.Info("Setting creds to " + addValue);
                        creds = addValue;
                        break;
                    case "toemails":
                        logger.Info("Setting 'To' recipients to " + addValue);
                        toEmails = addValue;
                        break;
                    case "fromemail":
                        logger.Info("Setting 'From' email address to " + addValue);
                        fromEmail = addValue;
                        break;
                }
            }
        }

        private static void sendMail(vType3EntriesWithBO asr, bool simulateOnly)
        {

            body = "Importer of Record: <importer>\r\nCustom Bond Number: <bondnumber>\r\nBond Amount: <bondamount>\r\nEffective Date: <effectivedate>\r\nBond Type: <bondtype>\r\n";
            body += "Surety Code: <suretycode>\r\nDate of Entry: <dateofentry>\r\nFiler Code: <filercode>\r\nEntry Number: <entrynumber>\r\nEntry Type: <entrytype>\r\n";
            body += "District/Port of Entry: <portofentry>\r\nValue: <value>\r\nEstimated Duty: <duty>\r\nEstimated Taxes: <taxes>\r\nEstimated Fees: <fees>\r\nEstimated ADD: <add>\r\nEstmated CVD: <cvd>\r\n";
            body += "Estimated Bonded ADD: <bondadd>\r\nEstmated Bonded CVD: <bondcvd>\r\n";
            body += "File Date: <filedate>\r\n";
            body += "\r\n\r\n";
            body += "The Customs Bond Desk is required to review this entry detail with the Customs Broker \r\nand alert IFIC Underwriting for the possible need to collateralize the Custom Bond.";

            bodyReplace("<importer>", asr.importer_of_record + " " + asr.importer_name);
            bodyReplace("<bondnumber>", asr.bond_number);
            bodyReplace("<bondamount>", String.Format("{0:C0}", asr.bond_liability_amount));
            if (asr.bond_effective_date == null)
            {
                bodyReplace("<effectivedate>", "no effective date");
            }
            else
            {
                bodyReplace("<effectivedate>", ((DateTime)asr.bond_effective_date).ToShortDateString());
            }
            bodyReplace("<bondtype>", asr.bond_type.ToString() + " " + getBondType(asr.bond_type.ToString()));
            bodyReplace("<suretycode>", asr.surety_code);
            bodyReplace("<dateofentry>", ((DateTime)asr.entry_date).ToShortDateString());
            bodyReplace("<filercode>", asr.filer_code1 + " " + getFilerCode(asr.filer_code1));
            bodyReplace("<entrynumber>", asr.entry_number1);
            bodyReplace("<entrytype>", asr.entry_type + " " + getEntryType(asr.entry_type.ToString()));
            bodyReplace("<portofentry>", asr.district_port_of_entry1.ToString() + " " + getPortName(asr.district_port_of_entry1.ToString()));
            bodyReplace("<value>", String.Format("{0:C0}", asr.value));
            bodyReplace("<duty>", String.Format("{0:C}", asr.estimated_duty));
            bodyReplace("<taxes>", String.Format("{0:C}", asr.estimated_taxes));
            bodyReplace("<fees>", String.Format("{0:C}", asr.estimated_fee));
            bodyReplace("<add>", String.Format("{0:C}", asr.estimated_antidumping_duty));
            bodyReplace("<cvd>", String.Format("{0:C}", asr.estimated_countervailing_duty));
            bodyReplace("<bondadd>", String.Format("{0:C}", asr.bonded_antidumping_duty));
            bodyReplace("<bondcvd>", String.Format("{0:C}", asr.bonded_countervailing_duty));
            bodyReplace("<filedate>", String.Format("{0:C}", asr.file_date.ToShortDateString()));


            if (!simulateOnly)
            {
                SmtpClient client = new SmtpClient(smtpServer);
                var credArray = creds.Split('|');
                client.Credentials = new NetworkCredential(credArray[0], credArray[1]);
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                MailAddress from = new MailAddress(fromEmail);
                MailMessage msg = new MailMessage();
                msg.From = from;
                foreach (string s in toEmails.Split(','))
                {
                    msg.To.Add(s);
                }

                msg.Subject = "Alert: An ADD/CVD entry has been posted against a Custom Bond. Principal: " + asr.importer_name + " TIN: " + asr.importer_of_record + " Entry: " + asr.entry_number1;
                msg.Body = body;
                msg.BodyEncoding = System.Text.Encoding.UTF8;
                msg.SubjectEncoding = System.Text.Encoding.UTF8;
                client.SendCompleted += new SendCompletedEventHandler(SendCompletedCallback);
                string userState = "type 3 message";
                logger.Info("Sending message:");
                logger.Info(body);
                client.SendAsync(msg, userState);
            }
            else
            {
                logger.Info("simulateOnly is TRUE. Not sending email, only logging what WOULD be sent.");
                logger.Info(body);
            }

        }

        private static string getBondType(string bondType)
        {
            string returnedBondType = "Unknown";
            try
            {
                int bondTypeID = int.Parse(bondType);
                cbpmqdbEntities1 db = new cbpmqdbEntities1();
                var q = from p in db.BondTypes where p.BondTypeId == bondTypeID select p;
                return q.First().Name;
            }
            catch (Exception)
            {
                return returnedBondType;
            }
        }

        private static string getEntryType(string entryType)
        {
            string returnedEntryType = "Unknown";
            try
            {
                int entryTypeID = int.Parse(entryType);
                cbpmqdbEntities1 db = new cbpmqdbEntities1();
                var q = from p in db.EntryTypes where p.EntryTypeId == entryTypeID select p;
                return q.First().Name.Replace(')', ' ');
            }
            catch (Exception)
            {
                return returnedEntryType;
            }
        }
        private static string getFilerCode(string filerCode)
        {
            string returnedFilerCode = "Unknown";
            try
            {
                cbpmqdbEntities1 db = new cbpmqdbEntities1();
                var q = from p in db.FilerCodes where p.FilerCode1 == filerCode select p;
                return q.First().FilerName;
            }
            catch (Exception)
            {
                return returnedFilerCode;
            }
        }

        private static string getPortName(string portCode)
        {
            string returnedPortName = "Unknown";
            try
            {
                cbpmqdbEntities1 db = new cbpmqdbEntities1();
                decimal portCodeID = decimal.Parse(portCode);
                var q = from p in db.PortCodes where p.PortCodeId == portCodeID select p;
                return q.First().Name;
            }
            catch (Exception)
            {
                return returnedPortName;
            }
        }
        private static void bodyReplace(string tag, string textToUse)
        {
            if (textToUse == null) textToUse = string.Empty; ;
            body = body.Replace(tag, textToUse);
        }

        private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
            // Get the unique identifier for this asynchronous operation.
            String token = (string)e.UserState;

            if (e.Cancelled)
            {
                logger.Info("[{0}] Send canceled.", token);
            }
            if (e.Error != null)
            {
                logger.Error("[{0}] {1}", token, e.Error.ToString());
            }
            else
            {
                logger.Info("Message sent.");
            }
        }
    }
}
