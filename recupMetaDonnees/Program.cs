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
            i.GetSiteFolders("Compta");
            InstanceBot.ConvertToString(i.ListDesDossier).ForEach(Console.WriteLine);
            
            i.GetFolderContentTypes(nomListe);
            i.GetChampsDunContentType(nomContentType);
            i.SetContentTypeWithString(nomContentType);
            i.ToUploadFile();
            i.SetCollValue("Title", "Final test");
            i.SetCollValue("Coudioldt2", "1999");
           i.SetCollValue("Payeee", "true");
           
            Console.ReadLine();
        }
    }
}
            
