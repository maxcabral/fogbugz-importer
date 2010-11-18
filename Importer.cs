using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.IO;
using System.Net;
using System.Collections;
using System.Security.Permissions;
using Microsoft.Win32;

using OfficeOpenXml;

namespace FogBugzImporter
{
    class Importer
    {
        struct Attachment
        {
            public string name;
            public string nameOnDisk;
        }

        static char[] m_attachmentSeparator = new char[] { ';' };

        string m_url;
        string m_name;
        string m_password;
        string m_ticketsFile;
        string m_mediaDir;

        bool m_seenResolve;
        bool m_seenClose;

        string m_lastBug;
        string m_token;
        Dictionary<string, string> m_people = new Dictionary<string, string>();
        StringBuilder m_log;

        public Importer(string url, string name, string password, string ticketsFile, string mediaDir)
        {
            m_url = url;
            m_name = name;
            m_password = password;
            m_ticketsFile = ticketsFile;
            m_mediaDir = mediaDir;
        }

        public void Connect()
        {
            Dictionary<string, string> args = new Dictionary<string, string>();
            args.Add("cmd", "logon");
            args.Add("email", m_name);
            args.Add("password", m_password);

            string result = CallRESTAPIFiles(m_url, args, null);

            XmlTextReader reader = new XmlTextReader(new StringReader(result));
            XPathDocument doc = new XPathDocument(reader);
            XPathNavigator nav = doc.CreateNavigator();

            m_token = nav.Evaluate("string(response/token)").ToString();

            loadPeople();
        }

        public void Import()
        {
            m_log = new StringBuilder();

            FileInfo existingFile = new FileInfo(m_ticketsFile);
            using (ExcelPackage pkg = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = pkg.Workbook.Worksheets[1];
                int row = 2;

                string cmd = getCellValue(worksheet, row, 1);
                while (!string.IsNullOrEmpty(cmd))
                {
                    string dt = getCellValue(worksheet, row, 2);
                    string sProject = getCellValue(worksheet, row, 3);
                    string sArea = getCellValue(worksheet, row, 4);
                    string sFixFor = getCellValue(worksheet, row, 5);
                    string ixPriority = getCellValue(worksheet, row, 6);
                    string sTitle = getCellValue(worksheet, row, 7);
                    string sEvent = getCellValue(worksheet, row, 8);
                    string sPersonAssignedTo = getCellValue(worksheet, row, 9);
                    string dtDue = getCellValue(worksheet, row, 10);
                    string reporter = getCellValue(worksheet, row, 11);
                    string attachments = getCellValue(worksheet, row, 12);

                    if (cmd == "close" && !m_seenResolve)
                        executeCommand("resolve", dt, sProject, sArea, sFixFor, ixPriority, sTitle, sEvent, sPersonAssignedTo, dtDue, reporter, attachments);

                    if (cmd == "reopen" && !m_seenClose)
                        cmd = "reactivate";

                    executeCommand(cmd, dt, sProject, sArea, sFixFor, ixPriority, sTitle, sEvent, sPersonAssignedTo, dtDue, reporter, attachments);

                    Console.WriteLine(string.Format("Processed {0} row", row));

                    row++;
                    cmd = getCellValue(worksheet, row, 1);
                }
            }

            File.WriteAllText("log.xml", m_log.ToString(), Encoding.Unicode);
        }

        private void executeCommand(string cmd, string dt, string sProject, string sArea, string sFixFor,
            string ixPriority, string sTitle, string sEvent, string sPersonAssignedTo, string dtDue,
            string reporter, string attachments)
        {
            if (cmd != "new" && string.IsNullOrEmpty(m_lastBug))
                throw new InvalidOperationException("Possibly invalid tickets file.");

            Dictionary<string, string> args = new Dictionary<string, string>();
            args.Add("cmd", cmd);
            args.Add("token", m_token);

            if (cmd != "new")
                args.Add("ixBug", m_lastBug);
            if (!string.IsNullOrEmpty(dt))
                args.Add("dt", dt);

            string personId = getPersonId(reporter);
            if (!string.IsNullOrEmpty(personId))
                args.Add("ixPersonEditedBy", personId);

            if (!string.IsNullOrEmpty(sTitle))
                args.Add("sTitle", sTitle);
            if (!string.IsNullOrEmpty(sProject))
                args.Add("sProject", sProject);
            if (!string.IsNullOrEmpty(sArea))
                args.Add("sArea", sArea);
            if (!string.IsNullOrEmpty(sFixFor))
                args.Add("sFixFor", sFixFor);
            if (!string.IsNullOrEmpty(sPersonAssignedTo))
                args.Add("sPersonAssignedTo", sPersonAssignedTo);
            if (!string.IsNullOrEmpty(ixPriority))
                args.Add("ixPriority", ixPriority);
            if (!string.IsNullOrEmpty(dtDue))
                args.Add("dtDue", dtDue);
            if (!string.IsNullOrEmpty(sEvent))
                args.Add("sEvent", sEvent);

            List<Attachment> files = getAttachments(attachments);
            ASCIIEncoding encoding = new ASCIIEncoding();
            Dictionary<string, byte[]>[] rgFiles = null;
            if (files.Count > 0)
            {
                rgFiles = new Dictionary<string, byte[]>[files.Count];
                for (int i = 0; i < files.Count; i++)
                {
                    rgFiles[i] = new Dictionary<string, byte[]>();
                    rgFiles[i]["name"] = encoding.GetBytes("File" + (i + 1).ToString());
                    rgFiles[i]["filename"] = encoding.GetBytes(files[i].name);
                    rgFiles[i]["contenttype"] = encoding.GetBytes(GetMIMEType(files[i].name));
                    FileStream fs = new FileStream(Path.Combine(m_mediaDir, files[i].nameOnDisk), FileMode.Open);
                    BinaryReader br = new BinaryReader(fs);
                    rgFiles[i]["data"] = br.ReadBytes((int)fs.Length);
                    fs.Close();
                }
                args.Add("nFileCount", files.Count.ToString());
            }

            string result = CallRESTAPIFiles(m_url, args, rgFiles);
            if (cmd == "new")
            {
                XmlTextReader reader = new XmlTextReader(new StringReader(result));
                XPathDocument doc = new XPathDocument(reader);
                XPathNavigator nav = doc.CreateNavigator();
                m_lastBug = nav.Evaluate("string(response/case/@ixBug)").ToString();
                m_seenResolve = false;
                m_seenClose = false;
            }
            else if (cmd == "resolve")
                m_seenResolve = true;
            else if (cmd == "close")
                m_seenClose = true;
            else if (cmd == "reopen" || cmd == "reactivate")
            {
                m_seenClose = false;
                m_seenResolve = false;
            }

            m_log.Append(result);
        }

        private List<Attachment> getAttachments(string attachments)
        {
            List<Attachment> result = new List<Attachment>();

            if (!string.IsNullOrEmpty(attachments))
            {
                string[] splitResult = attachments.Split(m_attachmentSeparator, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < splitResult.Length; )
                {
                    Attachment a = new Attachment();
                    a.name = splitResult[i++];
                    a.nameOnDisk = Path.GetFileName(splitResult[i++]);
                    result.Add(a);
                }
            }

            return result;
        }

        private string getPersonId(string person)
        {
            if (m_people.ContainsKey(person))
                return m_people[person];

            return null;
        }

        private string getCellValue(ExcelWorksheet worksheet, int row, int column)
        {
            object o = worksheet.Cells[row, column].Value;
            if (o == null)
                return null;

            return o.ToString();
        }

        private void loadPeople()
        {
            Dictionary<string, string> args = new Dictionary<string, string>();
            args.Add("cmd", "listPeople");
            args.Add("token", m_token);

            string result = CallRESTAPIFiles(m_url, args, null);
            XmlTextReader reader = new XmlTextReader(new StringReader(result));
            XPathDocument doc = new XPathDocument(reader);
            XPathNavigator nav = doc.CreateNavigator();

            string resultsTag = "response/people/person";
            string ixName = "ixPerson";
            string sName = "sFullName";
            
            XPathNodeIterator nl = (XPathNodeIterator)nav.Evaluate(resultsTag);
            foreach (System.Xml.XPath.XPathNavigator n in nl)
            {
                m_people.Add(
                    n.Evaluate("string(" + sName + ")").ToString(),
                    n.Evaluate("string(" + ixName + ")").ToString());
            }
        }

        //
        // CallRestAPIFiles submits an API request to the FogBugz api using the 
        // multipart/form-data submission method (so you can add files)
        // Don't forget to include nFileCount in your rgArgs collection if you are adding files.
        //
        private string CallRESTAPIFiles(string sURL, Dictionary<string, string> rgArgs, Dictionary<string, byte[]>[] rgFiles)
        {

            string sBoundaryString = getRandomString(30);
            string sBoundary = "--" + sBoundaryString;
            ASCIIEncoding encoding = new ASCIIEncoding();
            UTF8Encoding utf8encoding = new UTF8Encoding();
            HttpWebRequest http = (HttpWebRequest)HttpWebRequest.Create(sURL);
            http.Method = "POST";
            http.AllowWriteStreamBuffering = true;
            http.ContentType = "multipart/form-data; boundary=" + sBoundaryString;
            string vbCrLf = "\r\n";

            Queue parts = new Queue();

            //
            // add all the normal arguments
            //
            foreach (System.Collections.Generic.KeyValuePair<string, string> i in rgArgs)
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
            string sent = "";
            foreach (Byte[] part in parts)
            {
                stream.Write(part, 0, part.Length);
                sent += encoding.GetString(part);
            }
            stream.Close();
            //txtSent.Text = sent;

            //
            // read the result
            //
            Stream r = http.GetResponse().GetResponseStream();
            StreamReader reader = new StreamReader(r);
            string retValue = reader.ReadToEnd();
            //txtReceived.Text = retValue;
            reader.Close();

            return retValue;
        }

        private string getRandomString(int nLength)
        {
            string chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXTZabcdefghiklmnopqrstuvwxyz";
            string s = "";
            System.Random rand = new System.Random();
            for (int i = 0; i < nLength; i++)
            {
                int rnum = (int)Math.Floor((double)rand.Next(0, chars.Length - 1));
                s += chars.Substring(rnum, 1);
            }
            return s;
        }

        // From http://www.codeproject.com/dotnet/ContentType.asp
        private string GetMIMEType(string filepath)
        {
            RegistryPermission regPerm = new RegistryPermission(RegistryPermissionAccess.Read, "\\\\HKEY_CLASSES_ROOT");
            FileInfo fi = new FileInfo(filepath);
            RegistryKey classesRoot = Registry.ClassesRoot;
            string dotExt = fi.Extension.ToLower();
            RegistryKey typeKey = classesRoot.OpenSubKey("MIME\\Database\\Content Type");

            foreach (string keyname in typeKey.GetSubKeyNames())
            {
                RegistryKey curKey = classesRoot.OpenSubKey("MIME\\Database\\Content Type\\" + keyname);
                if (curKey.GetValue("Extension") != null && curKey.GetValue("Extension").ToString().ToLower() == dotExt)
                    return keyname;
            }
            return "";
        }
    }
}
