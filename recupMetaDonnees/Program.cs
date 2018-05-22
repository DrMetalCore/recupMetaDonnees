using System;
using System.Collections.Generic;
using System.Security;

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
            string titreFichier = "C:/Users/luka/source/repos/recupMetaDonnees/recupMetaDonnees/jeTest3.txt";
            string a = "partnR@xiome";
         
           InstanceBot i = new InstanceBot(URL, titreFichier, "collab.ext@axiome-solution.fr", "partenR@xiome");
            
            i.GetSiteFolders("bot1");
            i.GetFolderContentTypes("Documents");
            i.GetChampsDunContentType("Document");
            InstanceBot.ConvertToString(i.ListDesSites).ForEach(Console.WriteLine);
            Console.WriteLine("/////////////////////////");
            InstanceBot.ConvertToString(i.ListDesDossier).ForEach(Console.WriteLine);
            Console.WriteLine("/////////////////////////");
            InstanceBot.ConvertToString(i.ListDesContentType).ForEach(Console.WriteLine);
            Console.WriteLine("/////////////////////////");
            InstanceBot.ConvertToString(i.ListDesField).ForEach(Console.WriteLine);
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
            
