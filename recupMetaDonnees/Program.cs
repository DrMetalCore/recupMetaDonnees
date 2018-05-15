using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;


namespace recupMetaDonnees
{
    class Program
    {
        static void Main(string[] args)
        {
            string URL = "http://loca-mle-w16:81/sites/testBot";
            //string URL = "http://loca-fcn-sp16/sites";
            string nomListe = "Facture";
            string nomContentType = "Facture";
            string titreFichier = "C:/Users/luka/source/repos/recupMetaDonnees/recupMetaDonnees/jeTest3.txt";

            InstanceBot i = new InstanceBot(URL, titreFichier, "luka", "Axiomestage64","loca");
            
            // i.GetAllSiteCollections();
            /*
            HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create("http://loca-fcn-sp16/sites/proj/_api/search/query?querytext='contentclass:sts_site'");

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
                
                for (int ind = 0; ind < 100; ind++)
                {
                    Console.WriteLine(jobj["d"]["query"]["PrimaryQueryResult"]["RelevantResults"]["Table"]["Rows"]["results"][ind]["Cells"]["results"][6]["Value"]);
                }
                
                responseReader.Close();
                Console.ReadLine();


            }
            catch (Exception e)
            {
                Console.Out.WriteLine(e.Message); Console.ReadLine();
            }
            */
            
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
        }


    }
}
            
