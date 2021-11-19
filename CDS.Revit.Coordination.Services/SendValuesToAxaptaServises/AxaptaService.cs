using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CDS.Revit.Coordination.Services
{
    public class AxaptaService
    {
        public string HOST { get; set; } = "https://tstaxapi.cds.spb.ru/";
        public string LOGIN { get; set; } = "nevis";
        public string PASSWORD { get; set; } = "HPJoP/Y/33NPdTeITGd0WQ==";

        public AxaptaService(string host, string login, string password)
        {
            HOST = host;
            LOGIN = login;
            PASSWORD = password;
        }

        private T GetContentFromWebRequest<T>(HttpWebRequest http)
        {

            var responseValue = http.GetResponse();
            var stream = responseValue.GetResponseStream();
            var sr = new StreamReader(stream);
            var content = sr.ReadToEnd();
            return JsonConvert.DeserializeObject<T>(content);
        }
        private ObservableCollection<ElementClassifier> GetElementClassifiers(string category)
        {
            ObservableCollection<ElementClassifier> result = new ObservableCollection<ElementClassifier>();
            //get data foa access to axapta request
            try
            {
                using (var client = new WebClient())
                {
                    AccessAX accessAX = GetAccessAX();
                    if (accessAX != null)
                    {
                        //get json
                        string authorization = "Authorization:" + accessAX.token_type + " " + accessAX.access_token;

                        var baseAddress = HOST + "api/Navis/ClassifierCodeTable?type=" + category;

                        var http = (HttpWebRequest)WebRequest.Create(new Uri(baseAddress));
                        http.Headers.Add(authorization);
                        http.Method = "GET";
                        result = GetContentFromWebRequest<ObservableCollection<ElementClassifier>>(http);
                    }
                }
            }
            catch
            {
                
            }
            return result;
        }

        private AccessAX GetAccessAX()
        {
            AccessAX result = new AccessAX();
            
            using (var client = new WebClient())
            {

                var values = new NameValueCollection();
                values["grant_type"] = "password";
                values["username"] = LOGIN;
                values["password"] = PASSWORD;

                var response = client.UploadValues(HOST + "api/Account/token", values);

                string responseString = Encoding.Default.GetString(response);

                result = JsonConvert.DeserializeObject<AccessAX>(responseString);
            }

            return result;
        }
        public List<SectionClassifier> GetAllClassifiersSections()
        {
            List<SectionClassifier> result = new List<SectionClassifier>();
            //get data foa access to axapta request
            try
            {
                using (var client = new WebClient())
                {
                    AccessAX accessAX = GetAccessAX();
                    if (accessAX != null)
                    {
                        //get json
                        string authorization = "Authorization:" + accessAX.token_type + " " + accessAX.access_token;

                        var baseAddress = HOST + "api/Navis/ClassifierCodeTableType";

                        var http = (HttpWebRequest)WebRequest.Create(new Uri(baseAddress));
                        http.Headers.Add(authorization);
                        http.Method = "GET";
                        result = GetContentFromWebRequest<List<SectionClassifier>>(http);
                    }
                }
            }
            catch
            {
            }
            return result;
        }
        public Dictionary<string, ObservableCollection<ElementClassifier>> GetAllElementClassifiersDict(List<SectionClassifier> projectSections)
        {
            Dictionary<string, ObservableCollection<ElementClassifier>> classifierDict = new Dictionary<string, ObservableCollection<ElementClassifier>>();
            foreach (SectionClassifier sName in projectSections)
            {
                classifierDict[sName.CodeType] = GetElementClassifiers(sName.CodeType);
            }
            return classifierDict;

        }
        public Dictionary<string, string> GetWorksFromAxapta()
        {
            Dictionary<string, string> resultDictionary = new Dictionary<string, string>();


            return resultDictionary;
        }

        public string SendToAxapta<T>(List<T> worksToSend)
        {
            string jsonStr = JsonConvert.SerializeObject(worksToSend);
            string newStr = UnidecodeSharpFork.Unidecoder.Unidecode(jsonStr);
            byte[] bytes = Encoding.Default.GetBytes(newStr);

            var content = "accessAX = null";

            using (var client = new WebClient())
            {
                AccessAX accessAX = GetAccessAX();
                if (accessAX != null)
                {
                    string authorization = "Authorization:" + accessAX.token_type + " " + accessAX.access_token;

                    var baseAddress = HOST + "api/Navis/AddNavisData";

                    var http = (HttpWebRequest)WebRequest.Create(new Uri(baseAddress));
                    http.Headers.Add(authorization);
                    http.Accept = "application/json; charset=windows-1251";
                    http.ContentType = "application/json; charset=windows-1251";
                    http.Method = "POST";

                    Stream newStream = http.GetRequestStream();
                    newStream.Write(bytes, 0, bytes.Length);
                    newStream.Close();

                    var responseValue = http.GetResponse();

                    var stream = responseValue.GetResponseStream();
                    var sr = new StreamReader(stream);
                    content = sr.ReadToEnd();
                }
                return content;
            }
        }
    }
}
