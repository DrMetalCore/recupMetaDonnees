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
            string URL = "http://loca-mle-w16:81/sites/testBot/compta";
            string nomListe = "Facture";
            string nomContentType = "Facture";
            string titreFichier = "jeTest.txt";
            //Console.WriteLine("Enter your password.");
            //SecureString password = GetPassword();
            ClientContext clientContext = new ClientContext(URL);
            ListCollection collect = RecupDonneesSharePoint.getAllListes(clientContext);
            List<string> list =RecupDonneesSharePoint.convertToString(RecupDonneesSharePoint.getChampsDunContentType(clientContext,RecupDonneesSharePoint.getContentTypesDuneList(clientContext, nomListe), nomContentType));
            //list.ForEach(Console.WriteLine);
            list.ForEach(Console.WriteLine);
            Console.ReadLine();


        }
    }
}
            
