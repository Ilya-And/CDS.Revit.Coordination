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

namespace CDS.Revit.Coordination.Services.Axapta
{
    public class AxaptaService
    {
        public static string HOST { get; set; } = "https://tstaxapi.cds.spb.ru/";
        public static string LOGIN { get; set; } = "nevis";
        public static string PASSWORD { get; set; } = "HPJoP/Y/33NPdTeITGd0WQ==";

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

        /*Метод получения всех работ в связке с классификатором из Axapta в формате:    id - классификатор
                                                                                        Name - наименование
                                                                                        WorkCodeID - список работ*/
        private ObservableCollection<ElementClassifier> GetElementClassifiers(string category)
        {
            ObservableCollection<ElementClassifier> result = new ObservableCollection<ElementClassifier>();
            //get data for access to axapta request
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

        /*Метод получения всех разделов работ из Axapta: АР
                                                         КЖ
                                                         ОВиВК
                                                         ЭМ)*/
        private List<SectionClassifier> GetAllClassifiersSections()
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

        /*Метод получения всех работ из Axapta в формате:   Archive - находится ли работа в архиве
                                                            Name - наименование работы
                                                            ParentCodeId - родительский код работы
                                                            ProjWorkCodeId - код вида работ
                                                            UnitId - единицы измерения(м2, м3 и т.д.)*/
        private ObservableCollection<AxaptaWorkset> GetAllAxaptaWorksetsMethod()
        {
            var result = new ObservableCollection<AxaptaWorkset>();
            //get data foa access to axapta request
            
            using (var client = new WebClient())
            {
                AccessAX accessAX = GetAccessAX();
                if (accessAX != null)
                {
                    //get json
                    string authorization = "Authorization:" + accessAX.token_type + " " + accessAX.access_token;

                    var baseAddress = HOST + "api/Navis/ProjWorkTable";

                    var http = (HttpWebRequest)WebRequest.Create(new Uri(baseAddress));
                    http.Headers.Add(authorization);
                    http.Method = "GET";
                    result = GetContentFromWebRequest<ObservableCollection<AxaptaWorkset>>(http);
                }
                //var values = new NameValueCollection();
                //values["grant_type"] = "password";
                //values["username"] = LOGIN;
                //values["password"] = PASSWORD;

                //var response = client.UploadValues(HOST + "api/Account/token", values);

                //string responseString = Encoding.Default.GetString(response);

                //AccessAX accessAX = JsonConvert.DeserializeObject<AccessAX>(responseString);

                //get json
                //string authorization = "Authorization:" + accessAX.token_type + " " + accessAX.access_token;

                //HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(HOST + "api/Navis/ProjWorkTable");

                //req.Headers.Add(authorization);

                //HttpWebResponse resp = (HttpWebResponse)req.GetResponse();

                //using (StreamReader stream = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                //{
                //    responseString = stream.ReadToEnd();
                //}
                //result = JsonConvert.DeserializeObject<ObservableCollection<AxaptaWorkset>>(responseString);
                //result = GetAxaptaWorksetByGroups(result);
            }
            return result;
        }

        /*Метод получения словаря в формате: Раздел(АР, КЖ и т.д.) : Список работ в связке с классификатором в формате:     id - классификатор
                                                                                                                            Name - наименование
                                                                                                                            WorkCodeID - список работ*/
        private Dictionary<string, ObservableCollection<ElementClassifier>> GetAllElementClassifiersDict(List<SectionClassifier> projectSections)
        {
            Dictionary<string, ObservableCollection<ElementClassifier>> classifierDict = new Dictionary<string, ObservableCollection<ElementClassifier>>();
            foreach (SectionClassifier sName in projectSections)
            {
                classifierDict[sName.CodeType] = GetElementClassifiers(sName.CodeType);
            }
            return classifierDict;

        }

        /*Метод получения словаря с данными из Axapta в формате:  Классификатор : Список работ в формате -  Archive - находится ли работа в архиве
                                                                                                            Name - наименование работы
                                                                                                            ParentCodeId - родительский код работы
                                                                                                            ProjWorkCodeId - код вида работ
                                                                                                            UnitId - единицы измерения(м2, м3 и т.д.)*/
        public Dictionary<string, List<AxaptaWorkset>> GetWorksFromAxapta()
        {
            Dictionary<string, List<AxaptaWorkset>> resultDictionary = new Dictionary<string, List<AxaptaWorkset>>();
            var sectionsFromAxapta = GetAllClassifiersSections();
            var dictionaryElementClassifiers = GetAllElementClassifiersDict(sectionsFromAxapta);
            var allAxaptaWorksets = GetAllAxaptaWorksetsMethod();

            foreach(string section in dictionaryElementClassifiers.Keys)
            {
                foreach(ElementClassifier elementClassifier in dictionaryElementClassifiers[section])
                {
                    string classifire = elementClassifier.id;
                    var worksFromElementClassifier = elementClassifier.WorkCodeID;
                    foreach(string work in worksFromElementClassifier)
                    {
                        var axaptaWorkset = (from w in allAxaptaWorksets
                                             where w.ProjWorkCodeId == work
                                             select w).FirstOrDefault();

                        if(classifire != null && axaptaWorkset != null)
                        {
                            if (!resultDictionary.Keys.Contains(classifire))
                            {
                                var axaptaWorksetList = new List<AxaptaWorkset>();
                                axaptaWorksetList.Add(axaptaWorkset);
                                resultDictionary[classifire] = axaptaWorksetList;
                            }
                            else
                            {
                                if (!resultDictionary[classifire].Contains(axaptaWorkset))
                                {
                                    resultDictionary[classifire].Add(axaptaWorkset);
                                }
                            }
                        }
                    }
                }
            }

            return resultDictionary;
        }

        /*Метод отправки данных в Axapta
         */
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
