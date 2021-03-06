﻿using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using Newtonsoft.Json.Linq;

namespace recupMetaDonnees
{
    public class InstanceBot
    {
        private string Domaine { get; set; }
        private string FilePath { get; set; }
        private string NomDossier { get; set; }
        private string Login;
        private string Mdp;
        private string DomaineUser;
        
       
        private ContentType TypeDuFichier { get; set; }
        private ClientContext ClientCtx { get; set; }
        private ListItem Fichier { get; set; }


        public Dictionary<string, string> ListDesSiteCollections { get; set; }
        //public List<Web> ListDesSites { get; set; }
        public List<List> ListDesDossier { get; set; }
        public List<ContentType> ListDesContentType { get; set; }
        public List<Field> ListDesField { get; set; }

        /* Peux être inutile
        public List<string> ListDesSitesString { get; set; }
        List<string> ListDesDossierString;
        List<string> contentTypeCollString;
        List<string> fieldCollString;
        */

        public InstanceBot(string url, string chemin, string log, string pwd, string dom)
        {
            ClientCtx = new ClientContext(url);
            FilePath = chemin;
            string[] split = ClientCtx.Url.Split('/');
            Domaine = split[0] + "//" + split[2];
            Login = log;
            Mdp = pwd;
            DomaineUser = dom;

            ListDesSiteCollections = new Dictionary<string, string>();
            //ListDesSites = new List<Web>();
            ListDesDossier = new List<List>();
            ListDesContentType = new List<ContentType>();
            ListDesField = new List<Field>();

            GetAllSiteCollections(url);
            //GetAllSubWebs();

        }
        private void GetAllSiteCollections(string url)
        {

            HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(url+"/_api/search/query?querytext='contentclass:sts_site'&trimduplicates=false&rowlimit=100");

            endpointRequest.Method = "GET";
            endpointRequest.Accept = "application/json;odata=verbose";
            NetworkCredential cred = new NetworkCredential(Login, Mdp, DomaineUser);
            endpointRequest.Credentials = cred;
            HttpWebResponse endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();
            try
            {
                WebResponse webResponse = endpointRequest.GetResponse();
                Stream webStream = webResponse.GetResponseStream();
                StreamReader responseReader = new StreamReader(webStream);
                string response = responseReader.ReadToEnd();

                JObject jobj = JObject.Parse(response);

                for (int ind = 0; ind < jobj["d"]["query"]["PrimaryQueryResult"]["RelevantResults"]["Table"]["Rows"]["results"].Count(); ind++)
                {

                    string urlCollection = jobj["d"]["query"]["PrimaryQueryResult"]["RelevantResults"]["Table"]["Rows"]["results"][ind]["Cells"]["results"][6]["Value"].ToString();
                    string nomCollection = jobj["d"]["query"]["PrimaryQueryResult"]["RelevantResults"]["Table"]["Rows"]["results"][ind]["Cells"]["results"][3]["Value"].ToString();
                    if (urlCollection.Contains("loca-fcn-sp16/sites/") == true)
                    {
                        string[] split = urlCollection.Split('/');
                        ClientCtx = new ClientContext(Domaine + "/sites/" + split[4]);
                        using (ClientCtx = new ClientContext(ClientCtx.Url))
                        {
                            ClientCtx.Credentials = new NetworkCredential(Login, Mdp, DomaineUser);
                            Web rootWeb = ClientCtx.Site.RootWeb;
                            ClientCtx.Load(rootWeb);

                            BasePermissions bp = new BasePermissions();
                            bp.Set(PermissionKind.AddListItems);
                            //
                            ClientResult<bool> viewListItems = rootWeb.DoesUserHavePermissions(bp);
                            try
                            {
                                ClientCtx.ExecuteQuery();
                                if (viewListItems.Value)
                                {
                                    ListDesSiteCollections.Add(nomCollection, split[4]);
                                }

                            }
                            catch
                            {
                                Console.WriteLine("\n ---> You don't have access to site {0}\n   -- Exception: ");
                            }

                        }
                    }
                }

                responseReader.Close();

            }
            catch (Exception e)
            {
                Console.Out.WriteLine(e.Message); Console.ReadLine();
            }
        }
        /*
        private void GetAllSubWebs()
        {
            // Get the SharePoint web  
            Web web = ClientCtx.Web;
            ClientCtx.Load(web, website => website.Webs, website => website.Title);

            // Execute the query to the server  
            try
            {
                ClientCtx.ExecuteQuery();
            }
            catch
            {
                Console.WriteLine("Quelquechose s'est mal passé dans l'exéctution de la requete pour avoir les sites veuillez verifier l'url");
                Console.Read();
                System.Environment.Exit(-1);
            }

            // Loop through all the webs  
            foreach (Web subWeb in web.Webs)
            {
                string newpath = Domaine + subWeb.ServerRelativeUrl;
                ListDesSites.Add(subWeb);
                ClientCtx = new ClientContext(newpath);
                if (subWeb.Webs != null) GetAllSubWebs();
            }
        }
        */
        public void GetSiteFolders(string nomSite)
        {
            
            foreach (KeyValuePair<string,string> s in ListDesSiteCollections)
            {
                if (s.Key.Equals( nomSite, StringComparison.InvariantCultureIgnoreCase))
                {
                    //Update the client context with the selected site
                    ClientCtx = new ClientContext(Domaine + "/sites/" + s.Value);
                }
            }

            
                ListCollection listColl = ClientCtx.Web.Lists;
                
            
                // Execute query. 
                ClientCtx.Load(listColl, lists => lists.Include(testList => testList.Title,
                                                                    testList => testList.BaseTemplate));
                //try
                //{
                    ClientCtx.ExecuteQuery();
                /*}
                catch
                {
                    Console.WriteLine("Quelquechose s'est mal passé dans la récupération des dossier veuillez verifier le nom du site");
                    Task.Delay(4000);
                    System.Environment.Exit(-2);
                }*/

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
                        if(true)
                        {
                            if (f.Group == "Custom Columns" && f.FromBaseType == false) ListDesField.Add(f);
                            //if (f.Title == "Content Type") ListDesField.Add(f);
                        }
                    }
                }

            }

        }

        //Pour convertir la collection il faut mettre null dans les autres paramètres 

        public void SetCollValue(string nomColl, object valeur)
        {

            Field f = ListDesField.Find(field => field.Title == nomColl);
            //if (f == null) Console.WriteLine("Veuiller verifier le nom du champ");
             if (f.TypeAsString == "Boolean")
            {
                try
                {
                    Fichier[nomColl] = Convert.ToBoolean(valeur);
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
                    Fichier[nomColl] = Convert.ToInt32(valeur);
                }
                catch
                {
                    Console.WriteLine("L'entré n'était pas un nombre");
                }

            }
            else if (f.TypeAsString == "Text" )
            {
                try
                {
                    Fichier[nomColl] = valeur.ToString();
                }
                catch
                {
                    Console.WriteLine("L'entré n'était pas une chaine de caractère");
                }
            }
                

            Fichier.Update(); // important, rembeber changes

            try
            {
                ClientCtx.ExecuteQuery();
            }
            catch
            {
                Console.WriteLine("Quelquechose s'est mal passé dans la mofication de la valeur d'un champs veuillez verifier le le champs et la valeur");
                Console.Read();
                System.Environment.Exit(-5);
            }

        }

        public void ToUploadFile()
        {

            // Add the ListItem
            using (ClientCtx = new ClientContext(ClientCtx.Url))
            {
                ClientCtx.Credentials = new NetworkCredential(Login, Mdp, DomaineUser);
                if (NomDossier == "Documents") NomDossier = "Shared Documents";
                Folder folder = ClientCtx.Web.GetFolderByServerRelativeUrl(ClientCtx.Url + "/" + NomDossier);
                FileCreationInformation fci = new FileCreationInformation();
                try
                { 
                    fci.Content = System.IO.File.ReadAllBytes(FilePath);
                }
                catch
                {
                    Console.WriteLine("Quelquechose s'est mal passé dans le dépot du fichier veuillez verifier le chemin du fichier");
                    Console.Read();
                    System.Environment.Exit(-6);
                }
                string[] cut = FilePath.Split('/');
                fci.Url = cut.Last();
                fci.Overwrite = true;

                Microsoft.SharePoint.Client.File fileToUpload = folder.Files.Add(fci);
                ClientCtx.Load(fileToUpload);

                Fichier = fileToUpload.ListItemAllFields;
                ClientCtx.Load(Fichier);
                ClientCtx.ExecuteQuery();
                SetCollValue("Content Type", TypeDuFichier.Name);
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
                    if (field.Title !="Content Type") listARetourner.Add(field.Title);
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
                Dictionary <string, string> collectionConverti = (Dictionary<string, string>)collection;
                foreach (KeyValuePair<string,string> w in collectionConverti)
                {
                    listARetourner.Add(w.Key);
                }
            }

            return listARetourner;
        }

       
    }
}
