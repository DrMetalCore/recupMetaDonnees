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

            //// Get the content type collection for the list "Custom"
            //ContentTypeCollection contentTypeColl = clientContext.Web.Lists.GetByTitle(nomListe).ContentTypes;
            ContentTypeCollection contentTypeColl = RecupDonneesSharePoint.recupContentTypeDuneList(clientContext, nomListe);

            clientContext.Load(contentTypeColl);
            clientContext.ExecuteQuery();

            FieldCollection fC = RecupDonneesSharePoint.recupChampDunContentType(clientContext, contentTypeColl, nomContentType);
            //// Display the Content Type name
            
            foreach (Field f in fC)
            {
                Console.WriteLine(f.Title);
            }
            
            Console.ReadLine();
        }
    }
}
