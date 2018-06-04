using Microsoft.SharePoint.Client;
using System;


namespace recupMetaDonnees
{
    class Program
    {
        static void Main(string[] args)
        {
            //string URL = "http://loca-mle-w16/sites/testBot";
           string URL = "https://axiomesolution.sharepoint.com/sites/iutbot";
            string nomListe = "Facture";
            string nomContentType = "Facture";
            string titreFichier = "C:/Users/luka/source/repos/recupMetaDonnees/recupMetaDonnees/jeTest4.txt";
            string a = "partnR@xiome";
         
           InstanceBot i = new InstanceBot(URL, titreFichier, "collab.ext@axiome-solution.fr", "partenR@xiome");
           
            i.GetSiteFolders("Enseignants");
            i.GetFolderContentTypes("Documents");
            i.GetChampsDunContentType("Sujet de TP");
            /*InstanceBot.ConvertToString(i.ListDesSiteCollections).ForEach(Console.WriteLine);
            Console.WriteLine("/////////////////////////");*/
            InstanceBot.ConvertToString(i.ListDesSites).ForEach(Console.WriteLine);
            Console.WriteLine("/////////////////////////");
            InstanceBot.ConvertToString(i.ListDesDossier).ForEach(Console.WriteLine);
            Console.WriteLine("/////////////////////////");
            InstanceBot.ConvertToString(i.ListDesContentType).ForEach(Console.WriteLine);
            Console.WriteLine("/////////////////////////");
            InstanceBot.ConvertToString(i.ListDesField).ForEach(Console.WriteLine);
            foreach (var f in i.ListDesField)
            {
                //Console.WriteLine(f.FieldTypeKind.ToString());
                if (f.FieldTypeKind == Microsoft.SharePoint.Client.FieldType.Choice)
                {
                    FieldChoice c = i.ClientCtx.CastTo<FieldChoice>(f);
                    foreach (var choix in c.Choices)
                    {
                        Console.WriteLine(choix);
                    }
                }

            }
            /*
            foreach (var f in i.ListDesField)
            {
                Console.WriteLine(f.Title);
                Console.WriteLine(f.Group);
                Console.WriteLine(f.FromBaseType);
                Console.WriteLine(f.Hidden);
                Console.WriteLine(f.Indexed);
                Console.WriteLine("********************");
            }
            */
            Console.WriteLine("/////////////////////////");
            Console.ReadLine();
            /*
            foreach (var v in i.ListDesField)
            {
                
                Console.WriteLine(v.Title);
                Console.WriteLine("///////////////////"+v.Group);
            }
            Console.ReadLine();
            
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
            
