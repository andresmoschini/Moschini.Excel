using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moschini.Excel.OpenXml;
using System.IO;

namespace Moschini.Excel.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        [DeploymentItem("Files\\Companies.xlsx")]
        public void ReadTest()
        {
            IExcelConnectionFactory ExcelConnectionFactory = new OpenXmlExcelConnectionFactory();
            using (var connection = ExcelConnectionFactory.CreateConnection("Companies.xlsx"))
            using (var reader = connection.CreateReader("Sheet1"))
            {
                var titles = ColumnTitleCollection.ReadTitlesRow(reader);
                var columns = new
                {
                    CompanyName = titles.GetLocationByColumnTitle("Name", false),
                    CompanyUrl = titles.GetLocationByColumnTitle("Url", false),
                    TierName = titles.GetLocationByColumnTitle("Tier", false),
                    TrackName = titles.GetLocationByColumnTitle("Track", false),
                    MailingAddress_Country = titles.GetLocationByColumnTitle("MailingAddress Country", false),
                    MailingAddress_State = titles.GetLocationByColumnTitle("MailingAddress State", false),
                    MailingAddress_Province = titles.GetLocationByColumnTitle("MailingAddress Province", false),
                    MailingAddress_City = titles.GetLocationByColumnTitle("MailingAddress City", false),
                    MailingAddress_Zip = titles.GetLocationByColumnTitle("MailingAddress Zip", false),
                    MailingAddress_Street1 = titles.GetLocationByColumnTitle("MailingAddress Street1", false),
                    MailingAddress_Street2 = titles.GetLocationByColumnTitle("MailingAddress Street2", false),
                    PrimaryPhone_Number = titles.GetLocationByColumnTitle("PrimaryPhone Number", false),
                    PrimaryPhone_Extension = titles.GetLocationByColumnTitle("PrimaryPhone Extension", false)
                };

                int count = 0;
                while (reader.Read())
                {
                    count++;
                    Console.WriteLine(@"
    CompanyName: {0}
    CompanyUrl: {1}
    TierName: {2}
    TrackName: {3}
    MailingAddress_Country: {4}
    MailingAddress_State: {5}
    MailingAddress_Province: {6}
    MailingAddress_City: {7}
    MailingAddress_Zip: {8}
    MailingAddress_Street1: {9}
    MailingAddress_Street2: {10}
    PrimaryPhone_Number: {11}
    PrimaryPhone_Extension: {12}",
                        reader.GetValue<string>(columns.CompanyName),
                        reader.GetValue<string>(columns.CompanyUrl),
                        reader.GetValue<string>(columns.TierName),
                        reader.GetValue<string>(columns.TrackName),
                        reader.GetValue<string>(columns.MailingAddress_Country),
                        reader.GetValue<string>(columns.MailingAddress_State),
                        reader.GetValue<string>(columns.MailingAddress_Province),
                        reader.GetValue<string>(columns.MailingAddress_City),
                        reader.GetValue<string>(columns.MailingAddress_Zip),
                        reader.GetValue<string>(columns.MailingAddress_Street1),
                        reader.GetValue<string>(columns.MailingAddress_Street2),
                        reader.GetValue<string>(columns.PrimaryPhone_Number),
                        reader.GetValue<string>(columns.PrimaryPhone_Extension));
                }
                Assert.AreEqual(7, count);
            }
        }

        [TestMethod]
        [DeploymentItem("Files\\Companies.xlsx")]
        public void WriteTest()
        {
            IExcelConnectionFactory ExcelConnectionFactory = new OpenXmlExcelConnectionFactory();
            using (var connection = ExcelConnectionFactory.CreateConnection("Companies.xlsx"))
            {
                using (var writer = connection.CreateRandomWriter("Sheet1"))
                {
                    writer.Write(1, "A", "Hola");
                }
                using (var reader = connection.CreateReader("Sheet1"))
                {
                    reader.Read();
                    var result = reader.GetValue<string>("A");
                    Assert.AreEqual("Hola", result);
                }
            }
        }
    }
}
