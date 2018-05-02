using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace recupMetaDonnees
{
    class Program
    {
        static void Main(string[] args)
        {
            //// String Variable to store the siteURL
            string URL = "http://loca-mle-w16:81/sites/testBot/compta/";
            string nomListe = "Facture";
            string nomContentType = "Facture";
            string titreFichier = "jeTest.txt";
            
            //// Get the context for the SharePoint Site to access the data
            ClientContext clientContext = new ClientContext(URL);
            ListCollection lC = RecupDonneesSharePoint.getAllListes(clientContext);
            foreach(List l in lC)
            {
                Console.WriteLine(l.Title);
            }
            ContentTypeCollection cC = RecupDonneesSharePoint.getContentTypesDuneList(clientContext, nomListe);
            ContentType contentType = RecupDonneesSharePoint.StringToContentType(nomContentType, cC);
            ListItem item = RecupDonneesSharePoint.uploadFile(clientContext, titreFichier, nomListe, contentType);
            RecupDonneesSharePoint.setCollValue(clientContext, item, "Cout2", 999);

            Console.ReadLine();

        }
    }
}
            
