using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Linq;
using System.Net;


namespace recupMetaDonnees
{
    class Program
    {
        static void Main(string[] args)
        {
            //string URL = "http://loca-mle-w16/sites/testBot";
           string URL = "http://loca-fcn-sp16/sites/it";
            string nomListe = "Facture";
            string nomContentType = "Facture";
            string titreFichier = "C:/Users/luka/source/repos/recupMetaDonnees/recupMetaDonnees/jeTest3.txt";

           // InstanceBot i = new InstanceBot(URL, titreFichier, "luka", "Axiomestage64","loca");
            
            // i.GetAllSiteCollections();
            
            HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create("http://loca-fcn-sp16/sites/search/_api/search/query?querytext='contentclass:sts_site'&trimduplicates=false&rowlimit=100");

            endpointRequest.Method = "GET";
            endpointRequest.Accept = "application/json;odata=verbose";
            NetworkCredential cred = new System.Net.NetworkCredential("luka", "Axiomestage64", "LOCA");
            endpointRequest.Credentials = cred;
            HttpWebResponse endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();
            try
            {
                WebResponse webResponse = endpointRequest.GetResponse();
                Stream webStream = webResponse.GetResponseStream();
                StreamReader responseReader = new StreamReader(webStream);
                string response = responseReader.ReadToEnd();
                
                JObject jobj = JObject.Parse(response);
                
                for (int ind = 0; ind < jobj["d"]["query"]["PrimaryQueryResult"]["RelevantResults"]["Table"]["Rows"]["results"].Count(); ind++)
                {
                    string urlCollection = jobj["d"]["query"]["PrimaryQueryResult"]["RelevantResults"]["Table"]["Rows"]["results"][ind]["Cells"]["results"][6]["Value"].ToString();
                    if (urlCollection.Contains("loca-fcn-sp16/sites/")==true)
                    {
                        Console.WriteLine(urlCollection);
                    }
                }
                
                responseReader.Close();
                Console.ReadLine();


            }
            catch (Exception e)
            {
                Console.Out.WriteLine(e.Message); Console.ReadLine();
            }
            
            /*
            i.GetSiteFolders("Compta");
            InstanceBot.ConvertToString(i.ListDesSites).ForEach(Console.WriteLine);

            i.GetFolderContentTypes(nomListe);
            i.GetChampsDunContentType(nomContentType);
            InstanceBot.ConvertToString( i.ListDesField).ForEach(Console.WriteLine);
            
            i.SetContentTypeWithString(nomContentType);
            i.ToUploadFile();
            i.SetCollValue("Title", "Fidsnal test2 credi");
            i.SetCollValue("Cout2", "19598");
            i.SetCollValue("Payeee", "false");
           
            Console.ReadLine();
            */
        }


    }
}
            
