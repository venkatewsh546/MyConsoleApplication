using System;
using System.Collections.Generic;
using System.IO;
using System.Data.OleDb;
using System.Data;
using System.Threading;
using System.Data.SqlClient;
//using PdfSharp;
//using PdfSharp.Drawing;
//using PdfSharp.Pdf;
//using System.Drawing;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Reporting.WebForms;
using Microsoft.AnalysisServices.AdomdClient;
using Microsoft.AnalysisServices;
using System.Text.RegularExpressions;
using System.Xml.Schema;
using WinSCP;
using System.Diagnostics;
using System.Globalization;
using System.Web.Script.Serialization;
using System.Net.Mail;
using System.Net;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.Exchange.WebServices.Data;
using System.Security.Policy;
using System.Net.Http;
using System.Text;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Excel;

namespace ConsoleApplication2
{
    public class Program
    {
        // static string filepath = @"C:\Users\pothugun\Desktop\tempfiles\pcare-corrected\SHT-2017-03-24 v6_APJ.xlsx";

        static void Main(string[] args)
        {

            // pdfexportfromssis();

            //ssrsreportcall();

            //cubedata();  
            //jobschedule();

            //Xmlparsing();

            //missedxmlsparsing();

            // rannumcheck();

            //rptdate();

            //DataTable dt= GetFilesAndFolders(@"c:\");


            // regextest();

            // xmltest();          

            // runshscript();

            //downloadfile();

            //renamefiles();

            //excelfind();

            //string str=crypto.EncryptnDecrypt.EncryptData("test", "test");
            //  string str1 = crypto.EncryptnDecrypt.DecryptData(str, "test");


            // Jsontest();


            //  List<string> result = new List<string>(Regex.Split("qs;uifbwdv sduifwebvlwv bwlf wdlbw df wvev vvb rew", @"(?<=\G.{4})", RegexOptions.Singleline));

            // Smtptest();

            //Console.WriteLine(DateTime.Now.ToString("mmssfff"));

            //  HttpWebRequest request = WebRequest.Create("https://script.google.com/macros/u/0/s/AKfycbzImBks0cY1Uk6EAnFJ7qol8YNnGKeXUZZH1ssJCNegRIX3cZbb/exec") as HttpWebRequest;
            //  //optional
            //  HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            //  Stream stream = response.GetResponseStream();


            // googledrive();

            // Redamails();


            // GooglesheettestAsync();


            //Dm_ssas_hive_poc();

            //bool pageExits = false;
            //ServicePointManager.Expect100Continue = true;
            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
            //       | SecurityProtocolType.Tls11
            //       | SecurityProtocolType.Tls12
            //       | SecurityProtocolType.Ssl3;

            //try
            //{
            //    HttpWebRequest request = WebRequest.Create(@"https://stackoverflows.com/questions/924679/c-sharp-how-can-i-check-if-a-url-exists-is-valid") as HttpWebRequest;
            //    request.Method = "HEAD";
            //    HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            //    response.Close();

            //    if (response.StatusCode == HttpStatusCode.OK)
            //    {
            //        pageExits = true;
            //    }

            //}
            //catch (Exception ex)
            //{
            //    pageExits = false;
            //}


            // EncryptAndDecryptdata();

            ReadandEncryptData();



        }

        private static void ReadandEncryptData()
        {
            EncryptAndDecryptMethods edm = new EncryptAndDecryptMethods();
            //string data=File.ReadAllText(@"C:\Users\pothugun\OneDrive - Hewlett Packard Enterprise\personaldocs\fildata.txt");
            //File.WriteAllText(@"C:\Users\pothugun\OneDrive - Hewlett Packard Enterprise\personaldocs\encfildata.txt",edm.EncryptPassword(data, "83659"));

            String encdata = File.ReadAllText(@"C:\Users\pothugun\OneDrive - Hewlett Packard Enterprise\personaldocs\encfildata.txt");
            String descdata = edm.DecryptPassword(encdata, "83659");

        }

        private static void EncryptAndDecryptdata()
        {
            List<System.String> datacol = new List<System.String>();
            EncryptAndDecryptMethods edm = new EncryptAndDecryptMethods();

            String textfiledata = File.ReadAllText(@"C:\Users\pothugun\OneDrive - Hewlett Packard Enterprise\personaldocs\EncryptDecrypt.txt");
            FileData efd = JsonConvert.DeserializeObject<FileData>(textfiledata);

            FileData dfd = new FileData
            {
                Mydata = new Mydata()
            };
            dfd.Mydata.Unamepass = new List<UnamePass>();
            dfd.Mydata.Cardinfo = new List<CardInfo>();

            foreach(UnamePass uname in efd.Mydata.Unamepass)
            {
                dfd.Mydata.Unamepass.Add(
                    new UnamePass()
                    {
                        Source = uname.Source,
                        UserName = edm.DecryptPassword(uname.UserName, "83659"),
                        Password = edm.DecryptPassword(uname.Password, "83659")
                    });
                    
            }

            foreach(CardInfo ct in efd.Mydata.Cardinfo)
            {
                dfd.Mydata.Cardinfo.Add(
                    new CardInfo()
                    {
                        Source = ct.Source,
                        CardNo = edm.DecryptPassword(ct.CardNo, "83659"),
                        IFSCCODE = edm.DecryptPassword(ct.IFSCCODE, "83659"),
                        Validthrough = edm.DecryptPassword(ct.Validthrough, "83659"),
                        ValidFrom = edm.DecryptPassword(ct.ValidFrom, "83659"),
                        NameOnCard = edm.DecryptPassword(ct.NameOnCard, "83659"),
                        ThreeDSecureCode = edm.DecryptPassword(ct.ThreeDSecureCode, "83659"),
                        CVV = edm.DecryptPassword(ct.CVV, "83659"),
                        Notes = edm.DecryptPassword(ct.Notes, "83659")
                    });
            }

            File.WriteAllText(@"C:\Users\pothugun\OneDrive - Hewlett Packard Enterprise\personaldocs\EncryptDecrypt_fulldata.json", 
                JsonConvert.SerializeObject(dfd));

        }

        private static void Dm_ssas_hive_poc()
        {
            DataTable dttable = new DataTable();
            DataTable columns = new DataTable();
            SqlCommand scol;
            String coldata;
            StringBuilder str = new StringBuilder();

            SqlConnection scon = new SqlConnection("Server =gvs91925.houston.hpecorp.net,2048;Initial Catalog=ISEE_PUB;Integrated Security=SSPI;");
            scon.Open();

            SqlCommand scmd = new SqlCommand
            {
                Connection = scon,
                CommandText = "select distinct o.name tabname,s.name schemaname from sys.objects o join sys.schemas s " +
                "on o.schema_id=s.schema_id where type='U' and o.name like 'IRS%'"
            };
            dttable.Load(scmd.ExecuteReader());


            foreach(DataRow dr in dttable.Rows)
            {
                try
                {
                    //scol = new SqlCommand
                    //{
                    //    Connection = scon,
                    //    CommandText = "select c.name+',' from sys.columns c where c.object_id = object_id('" + dr[1] + "." + dr[0] + "')  FOR XML path('')"
                    //};
                    //coldata = scol.ExecuteScalar().ToString();
                    //scol.Dispose();

                    str.Append(Environment.NewLine);
                    str.Append(Environment.NewLine);
                    //str.AppendFormat("sqoop import -Dorg.apache.sqoop.splitter.allow_text_splitter=true --connect " +
                    //    "\"jdbc:sqlserver://gvs91925.houston.hpecorp.net:2048;databaseName=ISEE_PUB;user=DMPROWrite;" +
                    //    "password=pro1r@ryp@\\$\\$w1qr\" --query  \"select {0} from [Production].[{1}] where \\$CONDITIONS\" " +
                    //    "--split-by \"{2}\" --fields-terminated-by \",\"  --delete-target-dir  --target-dir '/user/Input/{3}'",
                    //    coldata.TrimEnd(','), dr[0], coldata.Substring(0, coldata.IndexOf(',')), dr[0]);
                    //str.Append(Environment.NewLine);
                    //str.Append(Environment.NewLine);

                    //File.AppendAllText(@"C:\Users\pothugun\Desktop\temp\scoopscript.txt", str.ToString());
                    //str.Clear();
                    

                    scol = new SqlCommand
                    {
                        Connection = scon,
                        CommandText = "select c.name+' '+t.name+'('+cast(c.max_length as varchar)+'),' " +
                    "from sys.columns c join sys.types t on c.user_type_id = t.user_type_id " +
                    " where c.object_id = object_id('" + dr[1] + "." + dr[0] + "')  FOR XML path('')"
                    };

                    coldata = scol.ExecuteScalar().ToString();
                    scol.Dispose();

                    //str.AppendFormat("CREATE external TABLE {0}_RAW({1})ROW FORMAT DELIMITED FIELDS TERMINATED BY ',' stored as textfile location '/user/Input/{2}", dr[0], coldata.TrimEnd(','), dr[0]);
                    //str.Append(Environment.NewLine);
                    //str.Append(Environment.NewLine);

                    //File.AppendAllText(@"C:\Users\pothugun\Desktop\temp\scoopscript.txt",str.ToString());
                    //str.Clear();

                    str.AppendFormat("create table {0}({1}) STORED as orc tblproperties(\"orc.compress\" = \"SNAPPY\", " +
                        "\"orc.stripe.size\" = \"671088640\", \"orc.row.index.stride\" = \"50000\", \"orc.create.index\" = \"true\") " +
                        ";", dr[0], coldata.TrimEnd(',').Replace("int(4)","int").Replace("date(3)","date").Replace("nvarchar","varchar"),dr[0]);
                    File.AppendAllText(@"C:\Users\pothugun\Desktop\temp\hivetable.txt", str.ToString());
                    str.Append(Environment.NewLine);
                    str.Append(Environment.NewLine);
                    str.Clear();


                    scol.Dispose();


                   // "create table irs_dim_calendar STORED as orc tblproperties ("orc.compress"="SNAPPY","orc.stripe.size"="671088640","orc.row.index.stride"="500","orc.create.index"="true") AS select * from irs_dim_calendar_raw;";
                }
                catch(Exception ex)
                {
                   
                }
            }

        }

        private static void GooglesheettestAsync()
        {
            HttpClient client = new HttpClient();

            var values = new Dictionary<string, string>
                        {
                           { "From", "48751" },
                           { "Message", "dfiygvauf" },
                           { "Source", "dfisdvsvygvauf" }
                        };

            var content = new FormUrlEncodedContent(values);

            var response = client.PostAsync("https://script.google.com/macros/s/AKfycbyB7uduq_ZqLIGk9YSz7pEwhhIjcoGO51lA8Qe_fvNqdJ6xCXaW/exec", content);
            response.Wait();
            
            client.Dispose();

            //WebClient webClient = new WebClient();
            //byte[] resByte;
            //string resString;
            //byte[] reqString;

            //try
            //{
            //    webClient.Headers["content-type"] = "application/json";
            //    reqString = Encoding.Default.GetBytes(JsonConvert.SerializeObject(new Jsondata { From= "From", Message="dfiygvauf", Source="difyvsdf" }, Formatting.Indented));
            //    resByte = webClient.UploadData("https://script.google.com/macros/s/AKfycbyB7uduq_ZqLIGk9YSz7pEwhhIjcoGO51lA8Qe_fvNqdJ6xCXaW/exec?From=985478&Message=uyewtvfdf&Source=ugcvsdgc",
            //        "Post", reqString);
            //    resString = Encoding.Default.GetString(resByte);
            //    Console.WriteLine(resString);
            //    webClient.Dispose();

            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine(e.Message);
            //}
        }


        private static void Redamails()
        {

            ExchangeService serviceInstance = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
            {
                Credentials = new NetworkCredential("venkatesh.pothuguntla@hpe.com", "")
            };
            //serviceInstance.TraceEnabled = true;
            //serviceInstance.TraceFlags = TraceFlags.All;
            serviceInstance.AutodiscoverUrl("venkatesh.pothuguntla@hpe.com", delegate
            {
                return true;
            });
            //serviceInstance.Url = myUri;

            //object o = serviceInstance.FindItems(WellKnownFolderName.Inbox, new ItemView(10));
            FindItemsResults<Item> findResults = serviceInstance.FindItems(WellKnownFolderName.Inbox, new ItemView(10));
            foreach (Item item in findResults.Items)
            {
                EmailMessage message = EmailMessage.Bind(serviceInstance, item.Id);

                //if (message.Subject.Contains(DateTime.Now.ToString("dd/mm/yyyy")))
                //{
                if (message.HasAttachments && message.Attachments[0] is FileAttachment)
                {
                    FileAttachment fileAttachment = message.Attachments[0] as FileAttachment;
                    //Change the below Path   
                    fileAttachment.Load(@"C:\Users\pothugun\Desktop\tempfiles\" + fileAttachment.Name);

                }
                // }

                var attach = item.IsAttachment;
                Console.WriteLine(item.DisplayTo+"\n"+ item.Subject);
            }



        }

        private static void Googledrive()
        {
            UserCredential credential;
            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/drive-dotnet-quickstart.json");

                string[] Scopes = { DriveService.Scope.DriveReadonly };
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            string ApplicationName = "Drive API .NET Quickstart";
            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define parameters of request.
            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.PageSize = 10;
            listRequest.Fields = "nextPageToken, files(id, name)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                .Files;
            Console.WriteLine("Files:");
            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    Console.WriteLine("{0} ({1})", file.Name, file.Id);
                }
            }
            else
            {
                Console.WriteLine("No files found.");
            }
            Console.Read();     

    }

        private static void Smtptest()
        {
            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress("venkatesh546@live.com");
                    mail.To.Add("venkatesh546@gmail.com");
                    mail.Subject = "sending message from app";
                    mail.Body = "<body>" + "sending message from app" + " </body>";
                    mail.IsBodyHtml = true;
                    using (SmtpClient smtp = new SmtpClient("smtp.live.com", 25))
                    {
                        smtp.Credentials = new NetworkCredential("venkatesh546@live.com", "Venkatesh_546@Live");
                        smtp.Timeout = 0;
                        smtp.EnableSsl = true;
                        smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                        smtp.SendMailAsync(mail);
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.InnerException);
            }
        }

        private static void Jsontest()
        {
            string file = @"C:\Users\pothugun\Desktop\tempfiles\enc_data_text.json";
            //deserialize JSON from file  
            string Json = System.IO.File.ReadAllText(file);
            JavaScriptSerializer ser = new JavaScriptSerializer();
            var personlist = ser.Deserialize<FileData>(Json);

            List<UnamePass> unamePass = personlist.Mydata.Unamepass;
            List<CardInfo> cardinfo = personlist.Mydata.Cardinfo;
           
            //personlist.mydata.card_info.Add(new CardInfo()
            //{
            //    Source = "hdfs",
            //    notes = new Notes()
            //    {
            //        cardNo = "34864",
            //        ifsccode = "ahc",
            //        nameoncard = "qhdh ",
            //        securecode = "ygefb",
            //        validfrom = "kabca",
            //        validThrough = "bwjvc",
            //    }
            //});

           // personlist.mydata.uname_pass.Add(new UnamePass() { Source = "", UserName = "", Password = "" });


            //foreach(CardInfo cif in cardinfo)
            //{
            //    if(cif.Source== "hdfs")
            //    {
            //        cif.Source = "hdfc";
            //        cif.notes.cardNo = "45787985478455";
            //        cif.notes.ifsccode = "kjh548798";
            //        cif.notes.nameoncard = "venkates";

            //    }
            //}



            //System.IO.File.WriteAllText(file, ser.Serialize(personlist));
            

           

            //foreach (CardInfo ci in cardinfo)
            //{
            //    str.Add(ci.Source);
            //    str.Add(ci.notes.cardNo);
            //    str.Add(ci.notes.ifsccode);
            //    str.Add(ci.notes.nameoncard);
            //    str.Add(ci.notes.securecode);
            //    str.Add(ci.notes.validfrom);
            //    str.Add(ci.notes.validThrough);                

            //}


            //personlist.Add(new DataFile()
            //{        
            //Source = "test",
            //Password="test",
            //UserName="test"
            //});

            //var josonwr = ser.Serialize(personlist.ToArray());
            //System.IO.File.WriteAllText(file, josonwr);

            //var uname= (from d in personlist select d).Distinct();

            //string[] str = (from DataFile in personlist where DataFile.Source.Equals("SBI") select DataFile.Source).Distinct().ToArray();

            //foreach (DataFile str in personlist)
            //{
            //    string source = str.Source;
            //    string password = str.Password;
            //    string userName = str.UserName;

            //}



        }
#if false
        private static void Excelfind()
        {
            string excelfile = @"C:\Users\pothugun\Desktop\tempfiles\datalookup.xlsx";
            excel.Application app = new excel.Application();
            excel.Workbook wbs = app.Workbooks.Open(excelfile);

            try
            {              
                
                app.Visible = false;                
                excel.Worksheet ws = wbs.Worksheets[1] as excel.Worksheet;

                excel.Range currentFind = null;
              

                excel.Range Fruits = app.get_Range("B1", "G22");

                currentFind = Fruits.Find("venkatesh", Type.Missing,
            excel.XlFindLookIn.xlValues, excel.XlLookAt.xlPart,
            excel.XlSearchOrder.xlByRows, excel.XlSearchDirection.xlNext, false,
            Type.Missing, Type.Missing);

                currentFind.Address.ToString();

                while (!currentFind.Address.ToString().Equals("$G$22"))
                {
                    excel.Range rng = (excel.Range)ws.Cells[19, 19];

                    excel.Range row = rng.EntireRow;

                    row.Insert(excel.XlInsertShiftDirection.xlShiftDown, false);
                    wbs.Save();
                    Fruits = app.get_Range("B1", "G22");
                    currentFind = Fruits.Find("venkatesh", Type.Missing,
            excel.XlFindLookIn.xlValues, excel.XlLookAt.xlPart,
            excel.XlSearchOrder.xlByRows, excel.XlSearchDirection.xlNext, false,
            Type.Missing, Type.Missing);

                }
            }
            finally
            {         
            
            wbs.Close(false);
            app.Quit();
            }

        }

#endif

        private static void Renamefiles()
        {
            String[] files = Directory.GetFiles(@"C:\Users\pothugun\OneDrive - Hewlett Packard Enterprise\personaldocs\pay slips");
            string[] monthnames = DateTimeFormatInfo.CurrentInfo.MonthNames;

            foreach (string str in files)
            {
                string[] filesp = str.Split('\\');
                string[] filename = filesp[filesp.Length - 1].Split('-');
                string filenamenew = filename[0].Substring(0, 3);

                               
            foreach(string str1 in monthnames)
                {
                    if(str1.ToUpper().StartsWith(filenamenew.ToUpper()))
                    {
                        System.IO.File.Move(str, str.Replace(filename[0], str1));
                    }

                }

            }


        }

        private static void Downloadfile()
       {
            List<String> addre = new List<String>
            {
                "593C4:C6949fed-1139-422e-b124-a1ac01093439",
                "6f1c19c2-4479-4888-a510-d1b42f7a6eca",
                "a53496f0-2e0e-41c1-8969-5dfbf31d9689",
                "10194a31-a9d2-e711-a293-9c8e9966dbe2",
                "d0f3cb8b-afd2-e711-b096-941882354c23",
                "4582a5c3-89f6-4412-9599-142445490bdd",
                "a3074641-0d9f-4ab5-9c8e-53aa94e9b226",
                "ddcb641a-ad00-4fea-81b1-182ed48ad886",
                "9a1b89f4-5309-4934-9551-5f38afa3d8d9",
                "e683a645-fb43-49fc-a925-04418201645f",
                "e0bf24bd-c2b8-4f7b-a84f-9496ad719c19",
                "ea4251c6-8807-45ee-aa3f-2c4189f8b46d",
                "7038e188-be7e-429a-bd89-64f25921ffd8",
                "227bf259-2c32-40e1-975d-a5486d233912",
                "ea06929e-9ad2-e711-892a-1c98ec26448c",
                "48ff0854-4a75-4ad7-9c2a-2604a85872a1",
                "9dc18632-12c9-43e6-906d-ded27df6be69",
                "50cdbc5b-a4d2-e711-a472-6c3be5b85dae",
                "95f49043-4fdc-4724-ac7e-c9c72f771ebb",
                "b89aa50f-8f98-4003-97a9-bfb763d4f0e3",
                "e2c6dd2e-85ba-4b54-83b4-17b99d42d94f",
                "7e5c052b-13cd-43b7-8469-bf8ab66a7277",
                "29d0ce14-c4c5-43b0-b61e-85442cfafc6e",
                "01625d1f-82be-4d68-b70a-cb7210683a91",
                "37cb138f-dfb5-4f04-8f5e-eb5e316b25b0",
                "681f06d7-e4ce-4517-b6e4-e9c3eb3f7ba7",
                "49057430-4952-4042-ab0d-cdc09ef1389f",
                "a91a7a6d-49d8-455e-808e-6af0f8a01d77",
                "f6b52f02-16d8-48c2-bf49-f10b1e58921f",
                "55f12553-78bb-4aae-a2cf-2bc9c8bcecb5",
                "4dc5f6a1-1c05-4d3c-b58c-5a02f1a74efd",
                "cafbba65-a7f9-4594-bf23-f72c5458021f",
                "0dd6d6fb-d806-44d2-a7b6-fc1d1eba698e",
                "6f481377-bd7e-4d62-8a1f-39b308e27324",
                "4f150c09-d751-49a1-bc53-a9adba39e820",
                "c9ef0f7e-a47c-47b2-89b6-171d13aeb2e5",
                "f4695dbd-b4a5-4801-95ec-dffd28bea27a",
                "a3b94b25-afbd-41f1-849d-6136e4798796",
                "befd1303-a8f0-4eb2-9430-0dcbf12c1d63",
                "52b48e80-d1aa-4868-b12a-85df96ce8d97",
                "b549b4c1-131a-4ebe-92a2-125029132417",
                "9a3492ba-c788-4e28-b7e8-5556dc7c7045",
                "c33881fb-bb9d-42d5-a547-cbfd0a8c027e",
                "66e40689-da7a-4707-810f-558cad290884",
                "2a60d3f4-dc5e-4c19-9bfb-d5b15718a626",
                "f8e0db04-a615-4aac-a095-702a9c54eab2",
                "0421b804-8a23-4eb9-bf95-5b731608176f",
                "70aae343-080d-4b45-9a1f-4b8893f580ab",
                "f28b84d3-e7c8-4214-b362-777dd2caa608",
                "2a2e20c8-53ae-4939-9787-4480e09ae497",
                "467e7296-d6b5-4711-9537-d844ca96d5e9",
                "1e48ba77-0bab-4337-aa15-efc171996473",
                "43aaf868-0237-419f-bd17-55572f36f269",
                "5083a191-32b1-45dc-8d1e-7c9a0f1839a6",
                "f18612ed-8b83-4965-bf94-c7f16177813f",
                "5571bbb0-9618-4cd4-a9f2-6480918ee83e",
                "42372562-7f60-4f12-8c3a-99ffbd29156e",
                "26da5800-0cc0-4542-9837-11c69d73f287",
                "a440b1a3-5595-4bfe-ae04-c7204f1743d9",
                "81d4c21a-6db9-4b60-b187-3a19f131294f",
                "b238263d-1d47-449e-a0bd-bede580e5773",
                "6fd74f9b-f066-46b2-b991-31cc73cdfa12",
                "e15d1f44-a2ef-4bb9-b551-16ed9c763c7c",
                "828d0b9e-ff99-40e1-833b-93c7c6018058",
                "5845060e-f666-4e88-8c6c-0d0065c957d4",
                "fbe445e1-64f2-4126-b911-3970203c72c4",
                "1fa4489d-f326-4e08-8d1d-9ee59a47deb8",
                "aa283887-c201-4063-b870-fec4d196e6bc",
                "e58bf673-f90d-4e60-b085-7decaa9be44e",
                "19171735-9300-472b-b8c4-76c922dcaa4a",
                "bc9a3f1c-9b74-4884-9970-ad6555b6a9ea",
                "b8b7efd7-8e0a-4d59-8cc0-dd25d644f870",
                "611fd502-0e70-497c-8580-ddc0fa3fc520",
                "d2fae5ac-9ca5-4b0f-9a64-5a8415882749",
                "41a3e055-2400-4992-be42-69e783c5a156",
                "f5338db2-b3ac-4b69-a449-a7d4892fe13e",
                "9d3b71ad-16ee-4537-a8cf-881f5dc0bd25",
                "4e0ca147-4b80-47bb-b20f-5a039ee97ddb",
                "0fb90977-33b2-4360-920c-a3df888ce708",
                "c6d5ae80-6663-420e-8b06-44266322a727",
                "69421965-5409-416a-b3f1-6af6cccb5d60",
                "5ce59f57-a464-43af-b62e-b22c79e6d4ef",
                "0f90b0cc-5f3d-4004-9c2c-c6aa14442fc2",
                "a8d3e689-2773-46f6-a241-ead45d076347",
                "8016d02d-ce5e-46cb-8147-9b8218edac77",
                "ecacc67f-7640-4472-a506-5f4842e61aa4",
                "83f138eb-b241-4318-a739-03c73269bfcf",
                "9d99901a-f86b-4719-9624-9f973a59ed78",
                "58bfd6a3-fee1-489e-9f45-a616a204172d",
                "1e3ad918-8e3a-41fe-bad8-1458da56f397",
                "edafc966-c8a5-470c-86b9-fb7042991a14",
                "35ceecbd-1e34-4070-9fdd-cfb23a1a8f81",
                "05d1405c-c352-42ca-a0bd-c9c4e65055d6",
                "e779e8f9-7aad-4317-8b6c-2fe708f7558a",
                "3e3250c4-f5eb-42e2-88bc-a9c817315fe2",
                "954111b4-7e8c-4ddb-a59e-153777de5fb6",
                "25e99972-6f64-4ab1-93d1-d64b96b17291",
                "96476e99-68f3-4251-b71a-bd7342c38bdc",
                "47bbf2ef-89e4-4be9-b8e5-0169660c0ebf",
                "afded64f-0188-461f-8cea-75383642f4a2",
                "b9e37f01-2d1c-4db9-b458-cfe5cda27d7c",
                "f153a041-d437-45eb-9a02-e48d1e31f122"
            };

            foreach (string str in addre)
           {
          Process.Start("C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe","https://saui.corp.int.hpe.com/sadb-support/DownloadFile?filetype=all&filename=+"+str);
           }

       }
           

       private static void Runshscript()
       {
           // Set up session options
           SessionOptions sessionOptions = new SessionOptions
           {
               Protocol = Protocol.Sftp,
               HostName = "hdp-1.pl.sarpt1.g4ihos.itcs.hpecorp.net",
               UserName = "hos",
               SshHostKeyFingerprint = "ssh-rsa 2048 5a:38:d0:85:7e:42:85:f2:dd:36:79:ab:b4:f5:52:91",
               SshPrivateKeyPath = @"C:\Users\pothugun\OneDrive - Hewlett Packard Enterprise\hadoop\new sandbox.ppk",
               
           };

           using (Session session = new Session())
           {
               // Connect
               session.Open(sessionOptions);

               session.ExecuteCommand("cd /home/hos");
               var list = session.ExecuteCommand("ls");

           }

       }

       public static void ValidationCallBack(object sender, ValidationEventArgs e)
       {
           // do something with any errors
       }

       private static void Regextest()
       {
          string text;
           string VERSION = string.Empty;
           string NAME = string.Empty;
           string VENDOR = string.Empty;

           Match matches;
           DataTable dt = new DataTable();
           SqlConnection scon = new SqlConnection("Server=g9w9026.houston.hpecorp.net,20482;Initial Catalog=DM_COLLECTIONS;Integrated Security=SSPI;");
           scon.Open();

           SqlCommand scmd = new SqlCommand("select * from [DMTA].[USR_REGEX_REVISION]", scon);

           dt.Load(scmd.ExecuteReader());

           text = "Linux - Red Hat Enterprise Linux ES 3";

           foreach(DataRow dr in dt.Rows )
           {
              matches = Regex.Match(text, dr["OSNAMEREGEX"].ToString().Replace("[:digit]", "0-9").Replace("[:blank:]", "\\t").Replace("[:digit:]", "0-9"), RegexOptions.None);
             VERSION = matches.ToString();
             NAME = dr["OSNAME"].ToString();
             VENDOR = dr["OSVENDOR"].ToString();

               if(!string.IsNullOrEmpty(VERSION))
               {
                   Console.WriteLine("match regex is :- " + VERSION);
                   break;
               }


           }
         

           

       }

       private static DataTable GetFilesAndFolders(string virtualDirPath)
       {
           DataTable filesAndFolders = new DataTable();
           filesAndFolders.Columns.Add("Name");
           filesAndFolders.Columns.Add("Path");
           filesAndFolders.Columns.Add("Type");


           string rootPath = virtualDirPath;
           if (Directory.Exists(rootPath))
           {
               string[] directories = Directory.GetDirectories(rootPath);
               for (int i = 0; i < directories.Length; i++)
               {
                   DirectoryInfo drinfo = new DirectoryInfo(directories[i]);
                   DataRow dr = filesAndFolders.NewRow();
                   dr["Name"] = drinfo.Name;
                   dr["Path"] = drinfo.FullName;
                   dr["Type"] = "D";
                   filesAndFolders.Rows.Add(dr);
               }
               string[] files = Directory.GetFiles(rootPath);
               for (int i = 0; i < files.Length; i++)
               {
                   FileInfo fInfo = new FileInfo(files[i]);
                   DataRow dr = filesAndFolders.NewRow();
                   dr["Name"] = fInfo.Name;
                   dr["Path"] = fInfo.FullName;
                   dr["Type"] = "F";
                   filesAndFolders.Rows.Add(dr);
               }
               return filesAndFolders;
           }
           return null;
       }


       private static void Rptdate()
       {
           DateTime rptdate = DateTime.Now;

          
           DateTime olddt = Convert.ToDateTime("2017/10/30");

           string irsjobstatus = "C";

          
           if (!irsjobstatus.ToUpper().Equals("F"))
           {


               //last day based on db date
               var lastDayOfMonth = DateTime.DaysInMonth(olddt.Year, olddt.Month);
               //last date based on db date
               DateTime lastdateofmonth = new DateTime(olddt.Year, olddt.Month, lastDayOfMonth);
               //adding 7 days to current db date
               DateTime newdt = olddt.AddDays(7);
               // if current db rpt date is run for monthly then pick next MONDAY 
               if (olddt.Date == lastdateofmonth.Date)
               {
                   rptdate = olddt;

                   while (!rptdate.ToString("dddd").ToUpper().Equals("MONDAY"))
                   {
                       rptdate = rptdate.AddDays(1);
                   }


                   //if month end date and new rpt date difference is 2 days theh pick next week MONDAY

                   if ((olddt.AddDays(1).ToString("dddd").ToUpper() == "MONDAY") || (olddt.AddDays(2).ToString("dddd").ToUpper() == "MONDAY"))
                   {
                       rptdate = rptdate.AddDays(7);
                   }

               }
               // if old rpt date and new rpt date month is same then return new rpt date 
               else if (olddt.Month == newdt.Month)
               {
                   rptdate = newdt;

                   while (!rptdate.ToString("dddd").ToUpper().Equals("MONDAY"))
                   {
                       rptdate = rptdate.AddDays(1);
                   }

               }
               // new rpt date and old db rpt dates months are different then it's a monthly run
               else if (olddt.Month != newdt.Month)
               {

                   rptdate = new DateTime(olddt.Year, olddt.Month, lastDayOfMonth);
               }

               // if new rpt date and month end date difference is 2 days theh run for monthly 

               //if ((rptdate.AddDays(1).Date == lastdateofmonth.Date) || (rptdate.AddDays(2).Date == lastdateofmonth.Date))
               //{
               //    rptdate = new DateTime(olddt.Year, olddt.Month, lastDayOfMonth);
               //}
           }
           else
           {
               rptdate = olddt;
           }
           Console.WriteLine(rptdate);
       }

       private static void Rannumcheck()
       {
           Random rnd = new Random();
           Random rnd1 = new Random();
           Random rnd2 = new Random();
           int i=0;

           while (i <= 300)
           {
               int month = rnd.Next(1, 100);
               int dice = rnd1.Next(1, 100);
               int card = rnd2.Next(1, 100);

               Console.WriteLine("Month :" + month.ToString() + "   dice :" + dice.ToString() + "   card :" + card.ToString());
               
               
               i =i + 1;
           }

       }

       private static void Missedxmlsparsing()
       {
           string METRIC_FILE = string.Empty;
           string LAST_METRIC_SIGNAL = string.Empty;
           string CMS_GDID = string.Empty;
           string CMS_SN = string.Empty;
           string CMS_PN = string.Empty;

           string inserstmt = string.Empty;
           string insertval = string.Empty;

           DataTable columncheck;

           try
    {
        columncheck = new DataTable();
        columncheck.Columns.Add("columnnames");

      //  int maxcount = 1000000;
      //  int initmincount =1;
    
    SqlConnection rmccon = new SqlConnection("Server=g9w8878.houston.hpecorp.net,20480;Initial Catalog=ISEE_STG;Integrated Security=SSPI;");
    rmccon.Open();
                SqlCommand rmccmd = new SqlCommand
                {
                    Connection = rmccon,
                    CommandTimeout = 0,

                    CommandText = "select distinct pkg.seq_id from SADB_Metrics_Collection_Pkg8 pkg left join (select distinct seq_id  from SADB_Metrics_Collection_Pkg_XML_allvalues(nolock) where seq_id between 7000001 and 8000000) sq on pkg.seq_id=sq.seq_id where sq.seq_id is null"
                };
                SqlDataReader mid = rmccmd.ExecuteReader();

             
      // int mincount = Convert.ToInt32(rmccmd.ExecuteScalar()) ;

      // mincount = mincount + 1;

    //------running loop taking each row into consideration-------------
      // for (int i = mincount; i <= maxcount; i++)

       List<int> missedlist = new List<int>();
       int s = 0;

               while(mid.Read())
               {
                   missedlist.Add(mid.GetInt32(0));
                   s = s + 1;
               }

               mid.Dispose();

    foreach (int i in missedlist)
       {
           try
           {
               columncheck.Rows.Clear();
               inserstmt = string.Empty;
               insertval = string.Empty;
               string tabname = Convert.ToString(Convert.ToInt32(i.ToString().Substring(1, 1)) + 1);


               rmccmd.CommandText = "select  METRIC_FILE,sadb.rcv_ts AS LAST_METRIC_SIGNAL,gdid as CMS_GDID,ATDCT_SRL_NR_TX as CMS_SN,ATDCT_PROD_NR_TX as CMS_PN " +
                                   " from dbo.SADB_Metrics_Collection_Pkg_XML8(nolock) main left join dbo.SADB_Metrics_Collection_Pkg8 (nolock) sadb on main.seq_id=sadb.seq_id " +
                                   " where main.seq_id =" + i;

               SqlDataReader dbread = rmccmd.ExecuteReader();

               if (dbread.Read())
               {

                   METRIC_FILE = dbread.GetValue(0).ToString().Replace("&", "").Replace("<????>", "????");
                   LAST_METRIC_SIGNAL = dbread.GetValue(1).ToString();
                   CMS_GDID = dbread.GetValue(2).ToString().Replace("'","''");
                   CMS_SN = dbread.GetValue(3).ToString().Replace("'", "''");
                   CMS_PN = dbread.GetValue(4).ToString().Replace("'", "''");
               }

               dbread.Dispose();

               inserstmt = "SEQ_ID,LAST_METRIC_SIGNAL,CMS_GDID,CMS_SN,CMS_PN";
               insertval = i.ToString() + " as SEQ_ID,'" + LAST_METRIC_SIGNAL + "' as LAST_METRIC_SIGNAL,'" + CMS_GDID + "' as CMS_GDID ,'" + CMS_SN + "' as CMS_SN,'" + CMS_PN + "' as CMS_PN";



               if (string.IsNullOrEmpty(METRIC_FILE) == false)
               {
                   DataSet ds = new DataSet();
                   StringReader xmlread = new StringReader(METRIC_FILE);

                   ds.ReadXml(xmlread);


                   foreach (DataTable tab in ds.Tables)
                   {
                       if (tab.TableName.Contains("LDID"))
                       {
                           inserstmt += ",LDID_Text";
                           insertval += ",'" + tab.Rows[0]["LDID_Text"].ToString().Replace("'", "''") + "' as LDID_Text";
                           columncheck.Rows.Add("LDID_Text");
                       }

                       if (tab.TableName.Equals("Property"))
                       {
                           foreach (DataRow dr in tab.Rows)
                           {
                               //----------For Basic Collection-----------------
                               if ((dr[0].ToString().ToLower().StartsWith("service::") || dr[0].ToString().ToLower().Equals("active_status") || dr[0].ToString().ToLower().Equals("transport_enabled")) && !dr[0].ToString().ToLower().Contains("message"))
                               {
                                   inserstmt += "," + dr[0].ToString().Replace("::", "_").Replace(" ", "_").Replace(".", "_").Replace("'", "_").Replace("-", "_");
                                   insertval += ",'" + dr[1].ToString().Replace("'", "''") + "' as " + dr[0].ToString().Replace("::", "_").Replace(" ", "_").Replace(".", "_").Replace("'", "_").Replace("-", "_");

                                   columncheck.Rows.Add(dr[0].ToString().Replace("::", "_").Replace(" ", "_").Replace("'", "_").Replace("-", "_"));

                                   // BASIC_COLLECTION_FUNCTIONAL = dr[1].ToString();
                               }

                           }
                       }

                       if (tab.TableName.Contains("GDID"))
                       {
                           inserstmt += ",GDID_Text";
                           insertval += ",'" + tab.Rows[0]["GDID_Text"].ToString().Replace("'", "''") + "' as GDID_Text";
                           columncheck.Rows.Add("GDID_Text");

                           // ENDPOINT_GDID = tab.Rows[0]["GDID_Text"].ToString();
                       }

                       if (tab.TableName.Contains("HP_ISEEEntitlementParameters"))
                       {
                           if (tab.Columns.Contains("SerialNumber"))
                           {
                               inserstmt += ",SerialNumber";
                               insertval += ",'" + tab.Rows[0]["SerialNumber"].ToString().Replace("'", "''") + "' as SerialNumber";

                               columncheck.Rows.Add("SerialNumber");
                               //SERIALNUMBER = tab.Rows[0]["SerialNumber"].ToString();
                           }
                           if (tab.Columns.Contains("ProductNumber"))
                           {
                               inserstmt += ",ProductNumber";
                               insertval += ",'" + tab.Rows[0]["ProductNumber"].ToString().Replace("'", "''") + "' as ProductNumber";

                               columncheck.Rows.Add("ProductNumber");

                               //PRODUCTNUMBER = tab.Rows[0]["ProductNumber"].ToString();
                           }

                       }
                   }

                            SqlCommand command1 = new SqlCommand
                            {
                                Connection = rmccon,
                                CommandTimeout = 0
                            };


                            try
                   {
                       foreach (DataRow dr in columncheck.Rows)
                       {      
                           command1.CommandText =  "if not exists( select 1 from sys.columns where object_id=object_id('SADB_Metrics_Collection_Pkg_XML_allvalues') and name='" + dr[0].ToString() + "') begin alter table isee_stg.dbo.SADB_Metrics_Collection_Pkg_XML_allvalues add " + dr[0].ToString() + " nvarchar(max) null end  ";
                           command1.ExecuteNonQuery();
                       }
                   }
                   catch
                     {
                        throw;
                   }

                   command1.CommandText = "insert into dbo.SADB_Metrics_Collection_Pkg_XML_allvalues  (" + inserstmt + ") select " + insertval;
                   command1.ExecuteNonQuery();

                   command1.Dispose();

               }

           }

           catch (Exception ex)
           {
              Console.WriteLine(ex.Message);
           }
       }
    }
    catch (Exception ex)
    {
     Console.WriteLine(ex.Message);
    }
}

       private static void Xmlparsing()
       {
           Dictionary<string, string> dict = new Dictionary<string, string>();

           DataSet ds = new DataSet();
           ds.ReadXml(@"C:\Users\pothugun\Desktop\tempfiles\test.xml");

           foreach (DataTable tab in ds.Tables)
           {
               if (tab.TableName.Equals("Property"))
               {
                   foreach (DataRow dr in tab.Rows)
                   {


                       if ((dr[0].ToString().StartsWith("Service::") || dr[0].ToString().Equals("Active_Status") || dr[0].ToString().Equals("Transport_Enabled")) && !dr[0].ToString().Contains("Message"))
                       {
                           Console.WriteLine(dr[0].ToString() + ":-" + dr[1].ToString());
                       }


                       if (dr[0].ToString() == "Service::Server_Basic_Configuration_Collection::DisplayColumn")
                       {
                           string val = dr[1].ToString();


                       }
                       else if (dr[0].ToString() == "Service::Subscription Manager::DisplayColumn")
                       {
                           string val = dr[1].ToString();
                       }

                   }

                   //string val = dict["Service::Server_Basic_Configuration_Collection::DisplayColumn"];
                   //string val1 = dict["Service::Subscription Manager::DisplayColumn"];
                   //string val2 = dict["Service::Server_Basic_Configuration_Collection::functional"];



               }
           }

       }

       private static void Jobschedule()
       {
           List<Thread> dythreads = new List<Thread>();

           try
           {
               string[] FileNames = Directory.GetFiles(@"C:\SA-project-codebase\SSIS\sarjobs-ITG\Batch", "*.bat");

               foreach (string str in FileNames)
               {
                   //Thread t = new Thread(new ThreadStart(() => new callbatch(str, this.ResultCallBack).executebatch()));
                   //t.Start();
                   //dythreads.Add(t);

               }
           }

           catch
           {
                throw;
            }
       }

       private static void PrintValues(DataTable table, string label)
       {
           Console.WriteLine(label);
           foreach (DataRow row in table.Rows)
           {
               foreach (DataColumn column in table.Columns)
               {
                   Console.Write("\t{0}", row[column]);
               }
               Console.WriteLine();
           }
       }

       private static DataTable CreateTestTable(string tableName)
       {
           DataTable table = new DataTable(tableName);
            DataColumn column = new DataColumn("id", typeof(System.Int32))
            {
                AutoIncrement = true
            };
            table.Columns.Add(column);

           column = new DataColumn("item", typeof(System.String));
           table.Columns.Add(column);

           // Add ten rows.
           DataRow row;
           for (int i = 0; i <= 9; i++)
           {
               row = table.NewRow();
               row["item"] = "item " + i;
               table.Rows.Add(row);
           }

           table.AcceptChanges();
           return table;
       }


       private static  void Cubedata()
       {

            AdomdConnection conn = new AdomdConnection
            {
                ConnectionString = "Data Source=hc4w00493.itcs.hpecorp.net;Initial Catalog=IRS_SNPN;Integrated Security=SSPI;Format=Tabular;"
            };
            // string query = " SELECT NON EMPTY { [Measures].[Device Count] } ON COLUMNS, NON EMPTY { ([REPORT DATE].[DATA TIME ID].[DATA TIME ID].ALLMEMBERS * [Geo].[Geography Info].[Region].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION ON ROWS FROM [ReportingCube] ";

            Server svr=new Server();
           svr.Connect("hc4w00493.itcs.hpecorp.net");
           Database db = svr.Databases.FindByName("IRS_SNPN");
           MeasureGroup mg = db.Cubes.FindByName("Sa_Connectivity_Cube").MeasureGroups.FindByName("SA DEVICES");
           PartitionCollection existingParts = mg.Partitions;

           AggregationDesignCollection agg = mg.AggregationDesigns;

           foreach(AggregationDesign ag1 in agg)
           {

           }

       }

       private static void Ssrsreportcall()
       {
            ReportViewer MyReportViewer = new ReportViewer
            {
                ProcessingMode = Microsoft.Reporting.WebForms.ProcessingMode.Remote
            };
            MyReportViewer.ServerReport.ReportServerUrl = new Uri("http://pothugun8:8081/ReportServer"); // Report Server URL
           MyReportViewer.ServerReport.ReportPath = "/Report/Report1";                     // Report Name
           MyReportViewer.ServerReport.Refresh();

           Microsoft.Reporting.WebForms.ReportParameter[] reportParameterCollection = new Microsoft.Reporting.WebForms.ReportParameter[1];
            reportParameterCollection[0] = new Microsoft.Reporting.WebForms.ReportParameter
            {
                Name = "id"                                                         //Parameter Name
            };
            reportParameterCollection[0].Values.Add("1");
           reportParameterCollection[0].Values.Add("2");
           reportParameterCollection[0].Values.Add("3"); //Parameter Value
           MyReportViewer.ServerReport.SetParameters(reportParameterCollection);


            byte[] bytes = MyReportViewer.ServerReport.Render("PDF", null, out string mimeType, out string encoding, out string extension, out string[] streamids, out Warning[] warnings);

            //Creatr PDF file on disk
            string pdfPath = @"C:\Users\pothugun\Desktop\tempfiles\jobstatus." + extension;       // Path to export Report.

           System.IO.FileStream pdfFile = new System.IO.FileStream(pdfPath, System.IO.FileMode.Create);
           pdfFile.Write(bytes, 0, bytes.Length);
           pdfFile.Close();
         

       }

       private static void Pdfexportfromssis()
       {

           Document doc = new Document(iTextSharp.text.PageSize.LETTER, 10, 10, 42, 35);

           try
           {
               string pdfFilePath = @"C:\Users\pothugun\Desktop\tempfiles\myPdf.pdf";
               //Create Document class object and set its size to letter and give space left, right, Top, Bottom Margin

               PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(pdfFilePath, FileMode.Create));
               doc.Open();//Open Document to write
               iTextSharp.text.Font font8 = FontFactory.GetFont("ARIAL", 7);

               Paragraph paragraph = new Paragraph("Using ITextsharp I am going to show how to create simple table in PDF document ");

               DataTable dt = new DataTable();
               dt.Columns.Add("col1");
               dt.Columns.Add("col2");
               dt.Rows.Add("venkatesh", "sample");



               if (dt != null)
               {

                   //Craete instance of the pdf table and set the number of column in that table

                   PdfPTable PdfTable = new PdfPTable(dt.Columns.Count);

                   PdfPCell PdfPCell = null;
                   //Add Header of the pdf table

                   PdfPCell = new PdfPCell(new Phrase(new Chunk("ID", font8)));

                   PdfTable.AddCell(PdfPCell);

                   PdfPCell = new PdfPCell(new Phrase(new Chunk("Name", font8)));

                   PdfTable.AddCell(PdfPCell);

                   //How add the data from datatable to pdf table

                   for (int rows = 0; rows < dt.Rows.Count; rows++)
                   {
                       for (int column = 0; column < dt.Columns.Count; column++)
                       {
                           PdfPCell = new PdfPCell(new Phrase(new Chunk(dt.Rows[rows][column].ToString(), font8)));

                           PdfTable.AddCell(PdfPCell);

                       }

                   }

                   PdfTable.SpacingBefore = 15f; // Give some space after the text or it may overlap the table

                   doc.Add(paragraph);// add paragraph to the document

                   doc.Add(PdfTable); // add pdf table to the document
               }
           }


           finally
           {

               //Close document and writer

               doc.Close();
           }




           ////Open the Excel file in Read Mode using OpenXml.
           // using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
           // {
           //     //Read the first Sheet from Excel file.
           //     Sheet sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();

           //     //Get the Worksheet instance.
           //     Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;

           //     //Fetch all the rows present in the Worksheet.
           //     IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();

           //     //Create a new DataTable.
           //     DataTable dt = new DataTable();

           //     //Loop through the Worksheet rows.
           //     foreach (Row row in rows)
           //     {
           //         //Use the first row to add columns to DataTable.
           //         if (row.RowIndex.Value == 1)
           //         {
           //             foreach (Cell cell in row.Descendants<Cell>())
           //             {
           //                 dt.Columns.Add(GetValue(doc, cell));
           //             }
           //         }
           //         else
           //         {
           //             //Add rows to DataTable.
           //             dt.Rows.Add();
           //             int i = 0;
           //             foreach (Cell cell in row.Descendants<Cell>())
           //             {
           //                 dt.Rows[dt.Rows.Count - 1][i] = GetValue(doc, cell);
           //                 i++;
           //             }
           //         }
           //     }
           // }

           Console.ReadLine();
       }

       //private static string GetValue(SpreadsheetDocument doc, Cell cell)
       //{
       //    string value = cell.CellValue.InnerText;
       //    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
       //    {
       //        value= doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText.ToString();
              
       //    }
       //    return value;
       //}
            
#region breaking big list into small lists
           //List<string> testpls = new List<string>() { "34", "35", "1X", "1Y", "2M", "3S", "3V", "4A", "4F", "4Q", "6H", "8L", "FA", "FE", "HA", "I6", "I7", "J2", "LA", "LI", "LJ", "LK", "LL", "LM", "LN", "MV", "NT", "SD", "SI", "SY", "TN", "TQ", "TR", "UZ" };

           // // List<string> pls = new List<string>() { "34" };

           //int listcount = testpls.Count;

           //int nooflists = (int)Math.Ceiling((double)listcount / 10);
           //int initval=1;

           //List<string> objlist=new List<string>();


           //while (initval <= nooflists)
           //{
           //    foreach (string str in testpls)
           //    {
           //        objlist.Add(str);                 
                   
           //        if(objlist.Count==10 || testpls.Count==objlist.Count)
           //        {
           //            break;
           //        }

           //    }

           //    foreach (string s in objlist)
           //    {
           //        testpls.Remove(s);
           //    }


           //    objlist.Clear();
           //    initval++;
            //}

#endregion
          

           //try
           // {
           //     SqlConnection scon = new SqlConnection();
           //     scon.ConnectionString = "Server=gvs91855.houston.hpecorp.net,2048;Initial Catalog=ISEE_STG;Integrated Security=SSPI;";
           //     scon.Open();
               
           //    string sconstr = "Server=gvs91855.houston.hpecorp.net,2048;Initial Catalog=ISEE_STG;Integrated Security=SSPI;";

           //     SqlCommand cmd = new SqlCommand();
           //     cmd.Connection = scon;
           //     cmd.CommandText = "select top 1 * from dbo.CORE_IB_DATA_ISS";
           //     DataTable dt = new DataTable();
           //     dt.Load(cmd.ExecuteReader());
           //     string query = "select * from dbo.CORE_IB_DATA_ISS where Product_Line_Id='";
           //     scon.Close();

           //     List<string> pls = new List<string> { "3V", "4Q", "FE", "LA", "MV", "SI", "SY", "TN", "UZ" };
           //     List<Thread> ts = new List<Thread>();

           //     foreach (string pl in pls)
           //     {
           //         Thread t = new Thread(new ThreadStart(() => new dataload(sconstr, query, pl, dt).loaddata()));
           //         t.Start();
           //         ts.Add(t);
           //     }

           //     foreach (Thread t1 in ts)
           //     {
           //         t1.Join();
           //     }
                
           //     Console.ReadLine();
           // }
           //catch(Exception ex)
           // {
           //     Console.WriteLine(ex.Message);
           //     Console.ReadLine();
           // }

      // }
       /*
          public static void loaddata(SqlConnection vcon,string query,DataTable dt,string pl)
        {
            Console.WriteLine("Thread {0} is started", pl);

            SqlCommand vcmd = new SqlCommand();
            vcmd.Connection = vcon;
            vcmd.CommandText = query + pl + "'"; ;
            SqlDataReader reader = vcmd.ExecuteReader();

            string data = "";

            foreach (DataColumn dtc in dt.Columns)
            {
                data += "\"" + dtc.ColumnName + "\"" + ",";
            }
            File.AppendAllText(@"C:\Users\pothugun\Downloads\globalshare\coreibdata_" + pl + ".csv", data.TrimEnd(',') + System.Environment.NewLine);
            data = "";


            while (reader.Read())
            {

                for (int i = 0; i < reader.FieldCount; i++)
                {
                    data += "\"" + reader[i] + "\"" + ",";
                }

                File.AppendAllText(@"C:\Users\pothugun\Downloads\globalshare\coreibdata_" + pl + ".csv", data.TrimEnd(',') + System.Environment.NewLine);
                data = "";

            }
        }

        */   

           // List<Int32> dt = new List<Int32>();

           // dt.Add(1);
           // dt.Add(2);
           // dt.Add(3);
           // dt.Add(4);


           //foreach(Int32 str in dt)
           //{
           //    StartThread(str);
           //}



           
            //MailMessage msg = new MailMessage("venkatesh.p2@hpe.com", "niveditha.l@hpe.com", "test", "testing");
            //msg.IsBodyHtml = true;
            //new SmtpClient("smtp1.hpe.com").Send(msg);
           // File.Create(@"C:\Users\pothugun\Desktop\OneDrive\files\test1.xlsx");
           //// string constr = @"Provider=Microsoft.Ace.OLEDB.12.0;Data Source=C:\Users\pothugun\Desktop\OneDrive\files\test1.xlsx;Extended Properties=&quot;EXCEL 12.0 Xml;HDR=YES;IMEX=1&quot;";

           ////OleDbConnection excelcon = new OleDbConnection(constr);
           //// excelcon.Open();
           //// OleDbCommand excelcmd = new OleDbCommand();
           //// excelcmd.Connection = excelcon;
           //// excelcmd.CommandText = "CREATE TABLE `outdata` (`PromotionKey` INTEGER,`PromotionAlternateKey` INTEGER,`EnglishPromotionName` NVARCHAR(255),`SpanishPromotionName` NVARCHAR(255),`FrenchPromotionName` NVARCHAR(255),    `EnglishPromotionType` NVARCHAR(50),`SpanishPromotionType` NVARCHAR(50),`FrenchPromotionType` NVARCHAR(50),`EnglishPromotionCategory` NVARCHAR(50),`SpanishPromotionCategory` NVARCHAR(50),`FrenchPromotionCategory` NVARCHAR(50),`StartDate` DATETIME,`EndDate` DATETIME,`MinQty` INTEGER,`MaxQty` INTEGER )";
           //// excelcmd.ExecuteNonQuery();
           //// excelcon.Close();
                            

       public static void StartThread(Int32 s)
       {
           Thread t = new Thread(() => { Loop(s); });
           t.Start();
           t.Join();
       }


       public static void Loop(Int32 i)
   {
       while(i<100)
   {
           Console.WriteLine(i);
           Thread.Sleep(1000);

   }
   }

       private static void RMCdatacheck()
       {
           DataTable rmcdt = new DataTable();
           DataTable metdt = new DataTable();

           SqlConnection rmccon = new SqlConnection("Server=16.189.41.62;Initial Catalog=dbo.APPDB_PRO;User ID=DMRead;Password=DMRead");
           rmccon.Open();

            SqlCommand rmccmd = new SqlCommand
            {
                CommandText = "select convert(varchar(10),cast(RECORDCREATIONTIMEUTC as date),112) dt ,count(1) reccount from parsingIDCFP(nolock) group by cast(RECORDCREATIONTIMEUTC as date)",
                Connection = rmccon
            };

            // rmcdt.Load(rmccmd.ExecuteReader());

            SqlDataReader reader = rmccmd.ExecuteReader();

           Dictionary<string, string> rmcdic = new Dictionary<string, string>();

           while(reader.Read())
           {
               rmcdic.Add(reader[0].ToString(), reader[1].ToString());

           }





            //foreach(DataRow dr in rmcdt.Rows)
            //{
            //    rmcdic.Add(dr[0].ToString(), dr[1].ToString());
            //}




            OleDbConnection metcon = new OleDbConnection
            {
                ConnectionString = "Data Source=MTRCSDBI;User ID=DMTA;Provider=OraOLEDB.Oracle.1;Persist Security Info=True;Password=Mko09ijnbhu&;"
            };
            metcon.Open();

            OleDbCommand metcmd = new OleDbCommand
            {
                CommandText = "select to_char(RMC_RECORDCREATIONTIMEUTC,'YYYY-MM-DD') as dt,count(1) rowcount from \"DMTA\".\"DIM_EVENTCAPTION_DATA\" group by to_char(RMC_RECORDCREATIONTIMEUTC,'YYYY-MM-DD');",
                Connection = metcon
            };

            metdt.Load(metcmd.ExecuteReader());
       
       
       }

private static void Method1()
{
    OleDbConnection excelcon;
    SqlConnection sqlcon;
    sqlcon = new SqlConnection("Server=g9w8878.houston.hpecorp.net,20480;Initial Catalog=DM_CASE;Integrated Security=True;");
    sqlcon.Open();

            SqlCommand cmd = new SqlCommand
            {
                Connection = sqlcon,
                CommandText = "select name from sys.objects where type='U'"
            };

            DataTable columns = new DataTable();
    columns.Load(cmd.ExecuteReader());
    DataTable data = new DataTable();

    foreach (DataRow dr in columns.Rows)
    {
        string str = dr["name"].ToString();
        cmd.CommandText = "sp_spaceused " + str;
        data.Load(cmd.ExecuteReader());
    }


    foreach (DataRow dr1 in data.Rows)
    {
        Console.WriteLine("TableName:-" + dr1["name"].ToString() + " records :- " + dr1["rows"].ToString());
    }



    try
    {
        string constr = @"Provider=Microsoft.Ace.OLEDB.12.0;Data Source=C:\Users\pothugun\Downloads\DM_CSC_RM_CLOSED_CASE_Details.xlsx;Extended Properties='EXCEL 12.0;HDR=YES;TypeGuessRows=200;IMEX=1;';";

        excelcon = new OleDbConnection(constr);
        excelcon.Open();

                OleDbCommand excelcmd = new OleDbCommand
                {
                    Connection = excelcon,
                    CommandText = "select distinct [closed date timestamp] from [Closed_cases$A2:bj] where [case number] is not null"
                };

                //OleDbDataReader excelreader = excelcmd.ExecuteReader();


                DataTable dt1 = new DataTable();
        dt1.Load(excelcmd.ExecuteReader());
        sqlcon = new SqlConnection("Server=g9w8878.houston.hpecorp.net,20480;Initial Catalog=DM_CASE;Integrated Security=True;");
        sqlcon.Open();

                //SqlCommand cmd = new SqlCommand();
                //cmd.Connection = sqlcon;
                //cmd.CommandText = "select top 1 * from CASE_STG_DATA";

                //DataTable columns = new DataTable();
                //columns.Load(cmd.ExecuteReader());


                SqlBulkCopy bulkcopy = new SqlBulkCopy(sqlcon, SqlBulkCopyOptions.TableLock, null)
                {
                    BatchSize = 100000,
                    DestinationTableName = "[dbo].[CASE_STG_DATA]",
                    BulkCopyTimeout = 100000000
                };


                foreach (DataColumn dc in columns.Columns)
        {
            bulkcopy.ColumnMappings.Add("[" + dc.ColumnName + "]", "[" + dc.ColumnName + "]");

        }
        // bulkcopy.WriteToServer(excelreader);
        excelcon.Close();
        sqlcon.Close();
    }

    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);

    }





    //string[] shtfile = Directory.GetFiles(@"C:\Users\pothugun\Desktop\data files", "s*.xlsx", SearchOption.TopDirectoryOnly);

    //List<OleDbDataReader> readers = new List<OleDbDataReader>();
    //OleDbDataReader r1;

    //fileload f1 = new fileload(shtfile[0].ToString());
    //Thread t1 = new Thread(f1.reader);

    //fileload f2 = new fileload(shtfile[0].ToString());
    //Thread t2 = new Thread(f2.reader);

    //fileload f3 = new fileload(shtfile[0].ToString());
    //Thread t3 = new Thread(f3.reader);

    //t1.Start();
    //t2.Start();
    //t3.Start();

    //t1.Join();
    //t2.Join();
    //t3.Join();


    //fileload f4 = new fileload();

    //readers = f4.readers;
}


    }

      
   
}
