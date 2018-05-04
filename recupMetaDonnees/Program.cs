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

            i.getAllSubWebs();
            i.getFoldersSite("Compta");
            i.getFolderContentTypes(nomListe);
            i.getChampsDunContentType(nomContentType);

            InstanceBot.convertToString(i.listDesSites).ForEach(Console.WriteLine);
            Console.WriteLine("-----------------------");
            InstanceBot.convertToString(i.listDesDossier).ForEach(Console.WriteLine);
            Console.WriteLine("-----------------------");
            InstanceBot.convertToString(i.contentTypeColl).ForEach(Console.WriteLine);
            Console.WriteLine("-----------------------");
            InstanceBot.convertToString(i.fieldColl).ForEach(Console.WriteLine);

            i.setContentTypeWithString(nomContentType);
            i.uploadFile();
            i.setCollValue("Cout2", 955);

            Console.ReadLine();
        }
    }
}
            
