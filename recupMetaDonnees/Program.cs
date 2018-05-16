using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
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
           string URL = "http://loca-fcn-sp16/sites/proj";
            string nomListe = "Facture";
            string nomContentType = "Facture";
            string titreFichier = "C:/Users/luka/source/repos/recupMetaDonnees/recupMetaDonnees/jeTest3.txt";

           InstanceBot i = new InstanceBot(URL, titreFichier, "luka", "Axiomestage64","loca");
            InstanceBot.ConvertToString(i.ListDesSiteCollections).ForEach(Console.WriteLine);

            // i.GetAllSiteCollections();
            
            
            
            i.GetSiteFolders("it");
            Dictionary<string, string> s = new Dictionary<string, string>();
            Console.ReadLine();
            /*
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
            
