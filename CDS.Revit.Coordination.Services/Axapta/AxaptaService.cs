using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace CDS.Revit.Coordination.Services.Axapta
{
    public class AxaptaService
    {
        /// <summary>
        /// Универсальный метод получения данных из Axapta
        /// </summary>
        /// <typeparam name="T">Тип возвращаемых данных</typeparam>
        /// <param name="senderType">Тип категории отправки</param>
        /// <param name="url">Адрес</param>
        /// <param name="httpMethod"></param>
        /// <returns></returns>
        private T GetDataByUrl<T>(SenderType senderType, string url, HttpMethod httpMethod)
        {
            using (var client = new WebClient())
            {
                AccessAX accessAX = GetAccessAX(senderType);
                if (accessAX != null)
                {
                    //get json
                    string authorization = "Authorization:" + accessAX.token_type + " " + accessAX.access_token;

                    var baseAddress = url;

                    var http = (HttpWebRequest)WebRequest.Create(new Uri(baseAddress));
                    http.Headers.Add(authorization);
                    http.Method = httpMethod.Method.ToString();

                    var responseValue = http.GetResponse();
                    var stream = responseValue.GetResponseStream();
                    var sr = new StreamReader(stream);
                    var content = sr.ReadToEnd();

                    return JsonConvert.DeserializeObject<T>(content);
                }

                throw new Exception("Token is null");
            }
        }
        private T GetContentFromWebRequest<T>(HttpWebRequest http)
        {

            var responseValue = http.GetResponse();
            var stream = responseValue.GetResponseStream();
            var sr = new StreamReader(stream);
            var content = sr.ReadToEnd();
            return JsonConvert.DeserializeObject<T>(content);
        }

        /// <summary>
        /// Получение токена
        /// </summary>
        /// <param name="senderType"></param>
        /// <returns></returns>
        private AccessAX GetAccessAX(SenderType senderType)
        {

            AccessAX result = new AccessAX();
            
            using (var client = new WebClient())
            {

                var values = new NameValueCollection();
                values["grant_type"] = "password";
                values["username"] = senderType == SenderType.Material ? AxaptaLoginPassword.TokenMaterialLogin : AxaptaLoginPassword.TokenWorkLogin;
                values["password"] = senderType == SenderType.Material ? AxaptaLoginPassword.TokenMaterialPassword : AxaptaLoginPassword.TokenWorkPassword;

                var response = client.UploadValues(LinkBuilder.Token, values);

                string responseString = Encoding.Default.GetString(response);

                result = JsonConvert.DeserializeObject<AccessAX>(responseString);
            }

            return result;
        }

        /// <summary>
        /// Метод получения всех работ в связке с классификатором из Axapta в формате:    id - классификатор
        ///                                                                               Name - наименование
        ///                                                                               WorkCodeID - список работ
        /// </summary>
        /// <param name="category"></param>
        /// <returns></returns>
        private ObservableCollection<ElementClassifier> GetElementClassifiers(string category)
        {
            ObservableCollection<ElementClassifier> result = GetDataByUrl<ObservableCollection<ElementClassifier>>(SenderType.Work, LinkBuilder.GetClassifierLink(category), HttpMethod.Get);
           
            return result;
        }

        /// <summary>
        /// Метод получения всех разделов работ из Axapta: АР, КЖ, ОВиВК, ЭМ
        /// </summary>
        /// <returns></returns>
        private List<SectionClassifier> GetAllClassifiersSections()
        {
            List<SectionClassifier> result = GetDataByUrl<List<SectionClassifier>>(SenderType.Work, LinkBuilder.ClassifierType, HttpMethod.Get);

            return result;
        }

        /*Метод получения всех работ из Axapta в формате:   Archive - находится ли работа в архиве
                                                            Name - наименование работы
                                                            ParentCodeId - родительский код работы
                                                            ProjWorkCodeId - код вида работ
                                                            UnitId - единицы измерения(м2, м3 и т.д.)*/
        private ObservableCollection<AxaptaWorkset> GetAllAxaptaWorksetsMethod()
        {
            ObservableCollection<AxaptaWorkset> result = GetDataByUrl<ObservableCollection<AxaptaWorkset>>(SenderType.Work, LinkBuilder.ProjWorkTable, HttpMethod.Get);
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
        public string SendToAxapta<T>(ObservableCollection<T> worksToSend, SenderType type)
        {
            string jsonStr = JsonConvert.SerializeObject(worksToSend);
            string newStr = UnidecodeSharpFork.Unidecoder.Unidecode(jsonStr);
            byte[] bytes = Encoding.Default.GetBytes(newStr);

            var content = "accessAX = null";

            using (var client = new WebClient())
            {
                AccessAX accessAX = GetAccessAX(SenderType.Work);
                if (accessAX != null)
                {
                    string authorization = "Authorization:" + accessAX.token_type + " " + accessAX.access_token;

                    var baseAddress = String.Empty;
                    if (type == SenderType.Work) baseAddress = LinkBuilder.Work;
                    if (type == SenderType.Material) baseAddress = LinkBuilder.Material;

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
