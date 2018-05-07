using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
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
            string nomListe = "Facture";
            string nomContentType = "Facture";
            string titreFichier = "C:/Users/luka/source/repos/recupMetaDonnees/recupMetaDonnees/jeTest3.txt";

            InstanceBot i = new InstanceBot(URL, titreFichier);

            i.GetAllSubWebs();
            i.GetFoldersSite("Compta");
            i.GetFolderContentTypes(nomListe);
            i.GetChampsDunContentType(nomContentType);
            i.SetContentTypeWithString(nomContentType);
            InstanceBot.ConvertToString(i.ListDesField).ForEach(Console.WriteLine);
            i.ToUploadFile();
            i.SetCollValue("Title", "Encule");
            i.SetCollValue("Cout2", "666");
           i.SetCollValue("Payeee", "true");

            Console.ReadLine();
        }
    }
}
            
