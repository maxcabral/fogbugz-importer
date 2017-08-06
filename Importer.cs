using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Xml;
using System.Xml.XPath;

using MimeSharp;
using CsvHelper;
using System.Linq;
using System.Threading.Tasks;

namespace FogBugzImporter
{
	struct Attachment
	{
		public string name;
		public string nameOnDisk;
	}

    class Importer
    {
        string caseFile;
        string attachmentDirectory;

        Mime mimeDetector = new Mime();
        FogBugzDriver driver;

        /// <summary>
        /// Gets or sets the attachment separators.
        /// </summary>
        /// <value>The attachment separators.</value>
		public char[] AttachmentSeparators { get; set; } = { ';' };

        public Importer(FogBugzDriver driver, string caseFile, string attachmentDirectory)
        {
            this.driver = driver;
            this.caseFile = caseFile;
            this.attachmentDirectory = attachmentDirectory;
        }

        public async Task ImportAsync()
        {
            FileInfo existingFile = new FileInfo(caseFile);
            using (var fileReader = new StreamReader(existingFile.OpenRead()))
            {
                using (var csvReader = new CsvReader(fileReader))
                {
                    List<string> headers = null;
                    if (csvReader.ReadHeader())
                    {
                        headers = csvReader.FieldHeaders.ToList();
                    }
                    else
                    {
                        Console.WriteLine("Could not read headers. Exiting");
                        return;
                    }

                    var headerCount = headers.Count();
                    while (csvReader.Read())
                    {
                        var record = csvReader.CurrentRecord;

                        var fogBugzCase = new Dictionary<String, String>(headers.Zip(record, (k, v) => new KeyValuePair<String, String>(k, v))
                                                                .Where(kp => !String.IsNullOrWhiteSpace((kp.Value))));

						if (fogBugzCase.TryGetValue("ixPersonAssignedTo", out var assignee))
						{
							if (!int.TryParse(assignee, out var _))
							{
								fogBugzCase.Remove("ixPersonAssignedTo");
								string personId = driver.GetPersonId(assignee);
								if (!string.IsNullOrEmpty(personId))
								{
									fogBugzCase.Add("ixPersonAssignedTo", personId);
								}
								else
								{
									Console.WriteLine($"Unable to look up personId for {assignee}");
								}
							}
						}

						if (fogBugzCase.TryGetValue("ixPersonEditedBy", out var reporter))
						{
							if (!int.TryParse(reporter, out var _))
							{
								fogBugzCase.Remove("ixPersonEditedBy");
								string personId = driver.GetPersonId(reporter);
								if (!string.IsNullOrEmpty(personId))
								{
									fogBugzCase.Add("ixPersonEditedBy", personId);
								}
								else
								{
									Console.WriteLine($"Unable to look up personId for {reporter}");
								}
							}
						}

						List<Dictionary<string, byte[]>> attachmentData = null;
						if (fogBugzCase.TryGetValue("attachments", out var attachments))
						{
                            fogBugzCase.Remove("attachments");
							List<Attachment> files = GetAttachments(attachments);
							ASCIIEncoding encoding = new ASCIIEncoding();

							if (files.Count > 0)
							{
                                attachmentData = new List<Dictionary<string, byte[]>>(files.Count);
								for (int i = 0; i < files.Count; i++)
								{
									attachmentData[i] = new Dictionary<string, byte[]>();
									attachmentData[i]["name"] = encoding.GetBytes("File" + (i + 1).ToString());
									attachmentData[i]["filename"] = encoding.GetBytes(files[i].name);
									attachmentData[i]["contenttype"] = encoding.GetBytes(mimeDetector.Lookup(files[i].name));
                                    using (FileStream fs = new FileStream(Path.Combine(attachmentDirectory, files[i].nameOnDisk), FileMode.Open))
                                    {
                                        BinaryReader br = new BinaryReader(fs);
                                        attachmentData[i]["data"] = br.ReadBytes((int)fs.Length);                                        
                                    }
								}
                                fogBugzCase.Add("nFileCount", files.Count.ToString());
							}
						}

                        await driver.ExecuteCommandAsync(fogBugzCase, attachmentData);
                    }
                }
            }
        }

        List<Attachment> GetAttachments(string attachments)
		{
			List<Attachment> result = new List<Attachment>();

			if (!string.IsNullOrEmpty(attachments))
			{
				string[] splitResult = attachments.Split(AttachmentSeparators, StringSplitOptions.RemoveEmptyEntries);
				for (int i = 0; i < splitResult.Length;)
				{
					Attachment a = new Attachment();
					a.name = splitResult[i++];
					a.nameOnDisk = Path.GetFileName(splitResult[i++]);
					result.Add(a);
				}
			}

			return result;
		}


	}

    /// <summary>
    /// FogBugz API driver.
    /// </summary>
    class FogBugzDriver {

		string fogbugzServiceUri;
		string userEmail;
		string userPassword;
		string fogbugzToken;
		Dictionary<string, string> nameToPersonIdMap = new Dictionary<string, string>();

        /// <summary>
        /// Gets a value indicating whether this <see cref="T:FogBugzImporter.FogBugzDriver"/> has connected to FogBugz.
        /// </summary>
        /// <value><c>true</c> if it has connected; otherwise, <c>false</c>.</value>
        public bool HasConnected { get; private set; }

		public FogBugzDriver(string fogbugzServiceUri, string fogbugzUserEmail, string fogbugzUserPassword)
		{
			this.fogbugzServiceUri = fogbugzServiceUri;
			userEmail = fogbugzUserEmail;
			userPassword = fogbugzUserPassword;
		}

		public FogBugzDriver(string fogbugzServiceUri, string fogbugzApiToken)
		{
			this.fogbugzServiceUri = fogbugzServiceUri;
			fogbugzToken = fogbugzApiToken;
		}

		public async Task ConnectAsync()
		{
			if (String.IsNullOrWhiteSpace(fogbugzToken))
			{
				fogbugzToken = await RetrieveApiTokenAsync();
			}
            nameToPersonIdMap = await RetrievePeopleIdsAsync();
            HasConnected = true;
		}

        /// <summary>
        /// Executes the command.
        /// </summary>
        /// <returns>The BugzId of the case operated upon, when available.</returns>
        /// <param name="caseData">Case data.</param>
        /// <param name="attachments">Attachments.</param>
        public async Task<int?> ExecuteCommandAsync(Dictionary<String,String> caseData, List<Dictionary<string, byte[]>> attachments)
        {
            String cmd;
            if (!caseData.TryGetValue("cmd", out cmd))
            {
                throw new InvalidDataException("A cmd value must be provided");
            }

            caseData.Add("token", fogbugzToken);

            string result = await CallRESTAPIFiles(fogbugzServiceUri, caseData, attachments);
            switch (cmd)
            {
                case "new":
                case "edit":
                case "assign":
                case "resolve":
                case "reactivate":
                case "close":
                case "reopen":
                    XmlTextReader reader = new XmlTextReader(new StringReader(result));
                    XPathDocument doc = new XPathDocument(reader);
                    XPathNavigator nav = doc.CreateNavigator();
                    var bugzIdStr = nav.Evaluate("string(response/case/@ixBug)")?.ToString();
                    if (bugzIdStr != null && !String.IsNullOrWhiteSpace(bugzIdStr))
                    {
                        if (int.TryParse(bugzIdStr, out var bugzId)){
                            return bugzId;
                        }
                    }
                    return null;
                default:
                    return null;
            }
        }

        public string GetPersonId(string person)
        {
            if (nameToPersonIdMap.ContainsKey(person))
                return nameToPersonIdMap[person];

            return null;
        }

        async Task<Dictionary<string, string>> RetrievePeopleIdsAsync()
        {
            var nameToIdMap = new Dictionary<string, string>();
            var args = new Dictionary<string, string>();
            args.Add("cmd", "listPeople");
            args.Add("token", fogbugzToken);

            string result = await CallRESTAPIFiles(fogbugzServiceUri, args, null);
            XmlTextReader reader = new XmlTextReader(new StringReader(result));
            XPathDocument doc = new XPathDocument(reader);
            XPathNavigator nav = doc.CreateNavigator();

            string resultsTag = "response/people/person";
            string ixName = "ixPerson";
            string sName = "sFullName";

            XPathNodeIterator nl = (XPathNodeIterator)nav.Evaluate(resultsTag);
            foreach (System.Xml.XPath.XPathNavigator n in nl)
            {
                nameToIdMap.Add(
                    n.Evaluate("string(" + sName + ")").ToString(),
                    n.Evaluate("string(" + ixName + ")").ToString());
            }

            return nameToIdMap;
        }

        async Task<string> RetrieveApiTokenAsync()
        {
			Dictionary<string, string> args = new Dictionary<string, string>();
			args.Add("cmd", "logon");
			args.Add("email", userEmail);
			args.Add("password", userPassword);

			string result = await CallRESTAPIFiles(fogbugzServiceUri, args, null);

			XmlTextReader reader = new XmlTextReader(new StringReader(result));
			XPathDocument doc = new XPathDocument(reader);
			XPathNavigator nav = doc.CreateNavigator();

			return nav.Evaluate("string(response/token)").ToString();
		}

        //
        // CallRestAPIFiles submits an API request to the FogBugz api using the 
        // multipart/form-data submission method (so you can add files)
        // Don't forget to include nFileCount in your rgArgs collection if you are adding files.
        //
        async Task<string> CallRESTAPIFiles(string sURL, Dictionary<string, string> rgArgs, IEnumerable<Dictionary<string, byte[]>> rgFiles)
        {

            string sBoundaryString = GetRandomString(30);
            string sBoundary = "--" + sBoundaryString;
            ASCIIEncoding encoding = new ASCIIEncoding();
            UTF8Encoding utf8encoding = new UTF8Encoding();
            HttpWebRequest http = (HttpWebRequest)WebRequest.Create(sURL);
            http.Method = "POST";
            http.AllowWriteStreamBuffering = true;
            http.ContentType = "multipart/form-data; boundary=" + sBoundaryString;
            string vbCrLf = "\r\n";

            Queue parts = new Queue();

            //
            // add all the normal arguments
            //
            foreach (KeyValuePair<string, string> i in rgArgs)
            {
                parts.Enqueue(encoding.GetBytes(sBoundary + vbCrLf));
                parts.Enqueue(encoding.GetBytes("Content-Type: text/plain; charset=\"utf-8\"" + vbCrLf));
                parts.Enqueue(encoding.GetBytes("Content-Disposition: form-data; name=\"" + i.Key + "\"" + vbCrLf + vbCrLf));
                parts.Enqueue(utf8encoding.GetBytes(i.Value));
                parts.Enqueue(encoding.GetBytes(vbCrLf));
            }

            //
            // add all the files
            //
            if (rgFiles != null)
            {
                foreach (Dictionary<string, byte[]> j in rgFiles)
                {
                    parts.Enqueue(encoding.GetBytes(sBoundary + vbCrLf));
                    parts.Enqueue(encoding.GetBytes("Content-Disposition: form-data; name=\""));
                    parts.Enqueue(j["name"]);
                    parts.Enqueue(encoding.GetBytes("\"; filename=\""));
                    parts.Enqueue(j["filename"]);
                    parts.Enqueue(encoding.GetBytes("\"" + vbCrLf));
                    parts.Enqueue(encoding.GetBytes("Content-Transfer-Encoding: base64" + vbCrLf));
                    parts.Enqueue(encoding.GetBytes("Content-Type: "));
                    parts.Enqueue(j["contenttype"]);
                    parts.Enqueue(encoding.GetBytes(vbCrLf + vbCrLf));
                    parts.Enqueue(j["data"]);
                    parts.Enqueue(encoding.GetBytes(vbCrLf));
                }
            }

            parts.Enqueue(encoding.GetBytes(sBoundary + "--"));

            //
            // calculate the content length
            //
            int nContentLength = 0;
            foreach (Byte[] part in parts)
            {
                nContentLength += part.Length;
            }
            http.ContentLength = nContentLength;

            //
            // write the post
            //
            Stream stream = http.GetRequestStream();
            foreach (Byte[] part in parts)
            {
                await stream.WriteAsync(part, 0, part.Length);
            }
            stream.Close();

            //
            // read the result
            //
            string retValue;
            var response = await http.GetResponseAsync();
            using (Stream r = response.GetResponseStream()) {
                StreamReader reader = new StreamReader(r);
                retValue = await reader.ReadToEndAsync();
            }

            return retValue;
        }

        string GetRandomString(int length)
        {
            string chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXTZabcdefghiklmnopqrstuvwxyz";
            string s = "";
            System.Random rand = new System.Random();
            for (int i = 0; i < length; i++)
            {
                int rnum = (int)Math.Floor((double)rand.Next(0, chars.Length - 1));
                s += chars.Substring(rnum, 1);
            }
            return s;
        }
    }
}
