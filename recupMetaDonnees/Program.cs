using Microsoft.SharePoint.Client;
using System;
using recupMetaDonnees;
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
            string URL = "http://loca-mle-w16:81/sites/testBot/compta";
            string nomListe = "Facture";
            string nomContentType = "Facture";

            //// Get the context for the SharePoint Site to access the data
            ClientContext clientContext = new ClientContext(URL);

            // Get the content type collection for the list "Custom"
            ContentTypeCollection contentTypeColl = clientContext.Web.Lists.GetByTitle(nomListe).ContentTypes;
            //ContentTypeCollection contentTypeColl = RecupDonneesSharePoint.getContentTypesDuneList(clientContext, nomListe);
            ListCollection lC = RecupDonneesSharePoint.getAllListes(clientContext);
            clientContext.Load(lC);
            clientContext.ExecuteQuery();
            List<string> liste = RecupDonneesSharePoint.convertToString(lC);
            liste.ForEach(Console.WriteLine);
            Console.WriteLine("---------------");

            clientContext.Load(contentTypeColl);
            clientContext.ExecuteQuery();
            liste = RecupDonneesSharePoint.convertToString(null, contentTypeColl);

            liste.ForEach(Console.WriteLine);
            Console.WriteLine("---------------");

            FieldCollection fC = RecupDonneesSharePoint.getChampsDunContentType(clientContext, contentTypeColl, nomContentType);
            //// Display the Content Type name
            clientContext.Load(fC);
            clientContext.ExecuteQuery();
            liste = RecupDonneesSharePoint.convertToString(null, null, fC);

            liste.ForEach(Console.WriteLine);
            /*
            foreach (Field f in fC)
            {
                Console.WriteLine(f.Title);
            }
            */
            /*
            ListCollection lC = RecupDonneesSharePoint.getAllListes(clientContext);
            foreach (List l in  lC)
            {
                Console.WriteLine(l.Title);
            }
            */
            Console.ReadLine();
        }
    }
}
