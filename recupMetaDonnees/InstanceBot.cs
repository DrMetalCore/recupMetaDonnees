using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Security;
using Microsoft.Online.SharePoint.TenantAdministration;
using AngleSharp.Network.Default;

namespace recupMetaDonnees
{
    public class InstanceBot
    {
        private string Domaine { get; set; }
        private string FilePath { get; set; }
        private string NomDossier { get; set; }
        private string Url;
        private string Login;
        private SecureString Mdp;



        private ContentType TypeDuFichier { get; set; }
        public ClientContext ClientCtx { get; set; }
        private ListItem Fichier { get; set; }


        public List<string> ListDesSiteCollections { get; set; }
        public List<Web> ListDesSites { get; set; }
        public List<List> ListDesDossier { get; set; }
        public List<ContentType> ListDesContentType { get; set; }
        public Dictionary<Field, object> ListDesField { get; set; }

        /* Peux être inutile
        public List<string> ListDesSitesString { get; set; }
        List<string> ListDesDossierString;
        List<string> contentTypeCollString;
        List<string> fieldCollString;
        */

        public InstanceBot(string url, string chemin, string log, string pwd)
        {
            ClientCtx = new ClientContext(url);
            Url = url;
            FilePath = chemin;
            string[] split = ClientCtx.Url.Split('/');
            Domaine = split[0] + "//" + split[2];
            Login = log;

            SecureString passWord = new SecureString();
            foreach (char c in pwd.ToCharArray()) passWord.AppendChar(c);
            Mdp = passWord;


            ListDesSiteCollections = new List<string>();
            ListDesSites = new List<Web>();
            ListDesDossier = new List<List>();
            ListDesContentType = new List<ContentType>();
            ListDesField = new Dictionary<Field, object>();

            //GetAllSiteCollections();
            GetAllSubWebs();

        }
        private void GetAllSiteCollections()
        {
            using (ClientCtx = new ClientContext("https://axiomesolution-admin.sharepoint.com/"))
            {
                ClientCtx.Credentials = new SharePointOnlineCredentials(Login, Mdp);
                SPOSitePropertiesEnumerable spp = null;
                Tenant tenant = new Tenant(ClientCtx);
                int startIndex = 0;

                while (spp == null || spp.Count > 0)
                {
                    spp = tenant.GetSiteProperties(startIndex, true);
                    ClientCtx.Load(spp);
                    ClientCtx.ExecuteQuery();

                    foreach (SiteProperties sp in spp)
                    {
                        ListDesSiteCollections.Add(sp.Title);
                    }
                    startIndex++;
                }
            }
        }

        public void GetAllSubWebs()
        {

            using (ClientCtx)
            {
                ClientCtx.Credentials = new SharePointOnlineCredentials(Login, Mdp);
                // Get the SharePoint web  
                Web web = ClientCtx.Web;
                ClientCtx.Load(web, website => website.Webs);

                // Execute the query to the server  
                //try
                //{
                ClientCtx.ExecuteQuery();
                //}
                //catch
                /*{
                    Console.WriteLine("Quelquechose s'est mal passé dans l'exéctution de la requete pour avoir les sites veuillez verifier l'url");
                    Console.Read();
                    System.Environment.Exit(-1);
                }
                */

                // Loop through all the webs  
                foreach (Web subWeb in web.Webs)
                {
                    string newpath = Domaine + subWeb.ServerRelativeUrl;
                    ListDesSites.Add(subWeb);

                    ClientCtx = new ClientContext(newpath);
                    if (subWeb.Webs != null) GetAllSubWebs();

                }
            }

            ClientCtx = new ClientContext(Url);
        }

        public void GetSiteFolders(string nomSite)
        {
            
                /*
                foreach (string s in ListDesSiteCollections)
                {
                    if (s.Equals(nomSite, StringComparison.InvariantCultureIgnoreCase))
                    {
                        //Update the client context with the selected site
                        ClientCtx = new ClientContext(Domaine + "/sites/" + s);
                    }
                }
                */
                foreach (Web s in ListDesSites)
                {
                    if (s.Title.Equals(nomSite, StringComparison.InvariantCultureIgnoreCase))
                    {
                        //Update the client context with the selected site
                        ClientCtx = new ClientContext(s.Url);
                    }
                }

                ListCollection listColl = ClientCtx.Web.Lists;


                // Execute query. 
                ClientCtx.Load(listColl, lists => lists.Include(testList => testList.Title,
                                                                    testList => testList.BaseTemplate));
                try
                {
                ClientCtx.ExecuteQuery();
                }
                catch
                {
                ClientCtx.Credentials = new SharePointOnlineCredentials(Login, Mdp);
                ClientCtx.ExecuteQuery();
                 }

                foreach (List list in listColl)
                {
                    if (list.BaseTemplate == 101 && list.Title != "Site Assets") // id dossier
                    {
                        ListDesDossier.Add(list);
                    }

                }
            


        }

        public void GetFolderContentTypes(string nomListe)
        {
            using (ClientCtx)
            {
                ClientCtx.Credentials = new SharePointOnlineCredentials(Login, Mdp);
                // Get the content type collection for the list nomListe

                NomDossier = nomListe;
                ContentTypeCollection contentTypeColl = ClientCtx.Web.Lists.GetByTitle(NomDossier).ContentTypes;
                //Execute the reques
                ClientCtx.Load(contentTypeColl);

                //try
                //{
                ClientCtx.ExecuteQuery();
                //}
                //catch
                //{
                //  Console.WriteLine("Quelquechose s'est mal passé dans la récupération des content types veuillez verifier le nom du dossier");
                //Console.Read();
                //System.Environment.Exit(-3);
                // }

                foreach (ContentType c in contentTypeColl)
                {
                    ListDesContentType.Add(c);
                }
            }
        }

        public void GetChampsDunContentType(string nomContentType)
        {
            SetContentTypeWithString(nomContentType);
            //// Get the field  collection for the content type nomContentType contenu dans la collection contentTypeCollection
            foreach (ContentType ct in ListDesContentType)
            {
                if (ct == TypeDuFichier)
                {
                    //Recupération des champs 
                    FieldCollection fieldColl = ct.Fields;
                    //Execution de la requette
                    // ClientCtx.Credentials = new NetworkCredential();
                    ClientCtx.Load(fieldColl);
                    try
                    {
                        ClientCtx.ExecuteQuery();
                    }
                    catch
                    {
                        Console.WriteLine("Quelquechose s'est mal passé dans la récupération des champs  veuillez verifier le nom du content type");
                        Console.Read();
                        System.Environment.Exit(-4);
                    }


                    foreach (Field f in fieldColl)
                    {
                        if (true)
                        {
                            if (f.FromBaseType == false && f.InternalName == "ContentType")
                            {
                                List<string> choix = new List<string>();
                                if (f.TypeAsString=="Boolean")
                                {
                                    ListDesField.Add(f, true);
                                }
                                if (f.FieldTypeKind == Microsoft.SharePoint.Client.FieldType.Choice)
                                {
                                    FieldChoice c = ClientCtx.CastTo<FieldChoice>(f);
                                    foreach (var ch in c.Choices)
                                    {
                                        choix.Add(ch);
                                    }
                                    ListDesField.Add(f, choix);
                                }
                                else ListDesField.Add(f, null);
                            }
                            
                        }
                    }
                }

            }

        }

        //Pour convertir la collection il faut mettre null dans les autres paramètres 

        public void SetCollValue(string nomColl, object valeur)
        {
            using (ClientCtx)
            {
                Field f = null;
                foreach (var pair in ListDesField)
                {
                    if (pair.Key.Title==nomColl || pair.Key.InternalName==nomColl)
                    {
                        f = pair.Key;
                    }
                }
                
                //if (f == null) Console.WriteLine("Veuiller verifier le nom du champ");
                if (f.TypeAsString == "Boolean")
                {
                    try
                    {
                        Fichier[f.InternalName] = Convert.ToBoolean(valeur);
                    }
                    catch
                    {
                        Console.WriteLine("L'entré n'était pas un booleen");
                    }
                }
                else if (f.TypeAsString == "Number" || f.TypeAsString == "Currency")
                {
                    try
                    {
                        Fichier[f.InternalName] = Convert.ToInt32(valeur);
                    }
                    catch
                    {
                        Console.WriteLine("L'entré n'était pas un nombre");
                    }

                }
                else if (f.TypeAsString == "Text")
                {
                    try
                    {
                        Fichier[f.InternalName] = valeur.ToString();
                    }
                    catch
                    {
                        Console.WriteLine("L'entré n'était pas une chaine de caractère");
                    }
                }


                Fichier.Update(); // important, rembeber changes
                ClientCtx.ExecuteQuery();
            }

        }

        public void ToUploadFile()
        {

            // Add the ListItem
            using (ClientCtx)
            {
                ClientCtx.Credentials = new SharePointOnlineCredentials(Login, Mdp);
                //if (NomDossier == "Documents") NomDossier = "Shared Documents";
                if (NomDossier == "Documents") NomDossier = "Documents partages";
                Folder folder = ClientCtx.Web.GetFolderByServerRelativeUrl(ClientCtx.Url + "/" + NomDossier);
                FileCreationInformation fci = new FileCreationInformation();
                //try
                //{ 
                fci.Content = System.IO.File.ReadAllBytes(FilePath);
                /* }
                 catch
                 {
                     Console.WriteLine("Quelquechose s'est mal passé dans le dépot du fichier veuillez verifier le chemin du fichier");
                     Console.Read();
                     System.Environment.Exit(-6);
                 }
                 */
                string[] cut = FilePath.Split('/');
                fci.Url = cut.Last();
                fci.Overwrite = true;

                Microsoft.SharePoint.Client.File fileToUpload = folder.Files.Add(fci);
                ClientCtx.Load(fileToUpload);

                Fichier = fileToUpload.ListItemAllFields;
                ClientCtx.Load(Fichier);
                ClientCtx.ExecuteQuery();
                SetCollValue("ContentType", TypeDuFichier.Name);
                string[] titre = cut.Last().Split('.');
                SetCollValue("Title", titre.First());
                // Now invoke the server, just one time

            }

        }

        public void SetContentTypeWithString(string contentString)
        {
            foreach (ContentType contentType in ListDesContentType)
            {
                if (contentType.Name == contentString) TypeDuFichier = contentType;
            }

        }

        public static List<string> ConvertToString(object collection)
        {
            List<string> listARetourner = new List<string>();
            if (collection.GetType().ToString() == "System.Collections.Generic.List`1[Microsoft.SharePoint.Client.List]")
            {
                List<List> collectionConverti = (List<List>)collection;
                foreach (List list in collectionConverti)
                {
                    listARetourner.Add(list.Title);
                }
            }
            else if (collection.GetType().ToString() == "System.Collections.Generic.List`1[Microsoft.SharePoint.Client.ContentType]")
            {
                List<ContentType> collectionConverti = (List<ContentType>)collection;
                foreach (ContentType contentType in collectionConverti)
                {
                    listARetourner.Add(contentType.Name);
                }
            }
            else if (collection.GetType().ToString() == "System.Collections.Generic.List`1[Microsoft.SharePoint.Client.Field]")
            {
                List<Field> collectionConverti = (List<Field>)collection;
                foreach (Field field in collectionConverti)
                {
                    if (field.InternalName != "ContentType") listARetourner.Add(field.Title);
                }
            }
            else if (collection.GetType().ToString() == "System.Collections.Generic.List`1[Microsoft.SharePoint.Client.Web]")
            {
                List<Web> collectionConverti = (List<Web>)collection;
                foreach (Web w in collectionConverti)
                {
                    listARetourner.Add(w.Title);
                }
            }
            else if (collection.GetType().ToString() == "System.Collections.Generic.Dictionary`2[System.String,System.String]")
            {
                Dictionary<string, string> collectionConverti = (Dictionary<string, string>)collection;
                foreach (KeyValuePair<string, string> w in collectionConverti)
                {
                    listARetourner.Add(w.Key);
                }
            }

            return listARetourner;
        }


    }
}