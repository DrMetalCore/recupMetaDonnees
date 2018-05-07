using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace recupMetaDonnees
{
    public class InstanceBot
    {
        private string domaine { get; set; }
        private string filePath { get; set; }
        public string nomDossier { get; set; }
        public string nomChamp { get; set; }
        public object valeurChamp { get; set; }

        private Web site { get; set; }
        private ContentType typeDuFichier { get; set; }
        private ClientContext clientContext { get; set; }
        private ListItem fichier { get; set; }


        public List<Web> listDesSites { get; set; }
        public List<List> listDesDossier { get; set; }
        public List<ContentType> listDesContentType { get; set; }
        public List<Field> listDesField { get; set; }

        /* Peux être inutile
        public List<string> listDesSitesString { get; set; }
        List<string> listDesDossierString;
        List<string> contentTypeCollString;
        List<string> fieldCollString;
        */

        public InstanceBot(string url, string chemin)
        {
            clientContext = new ClientContext(url);
            filePath = chemin;
            string[] split = clientContext.Url.Split('/');
            domaine = split[0] + "//" + split[2];

            listDesSites = new List<Web>();
            listDesDossier = new List<List>();
            listDesContentType = new List<ContentType>();
            listDesField = new List<Field>();

            getAllSubWebs();

        }
        public void getAllSubWebs()
        {
            // Get the SharePoint web  
            Web web = clientContext.Web;
            clientContext.Load(web, website => website.Webs, website => website.Title);

            // Execute the query to the server  
            clientContext.ExecuteQuery();

            // Loop through all the webs  
            foreach (Web subWeb in web.Webs)
            {
                string newpath = domaine + subWeb.ServerRelativeUrl;
                listDesSites.Add(subWeb);
                clientContext = new ClientContext(newpath);
                if (subWeb.Webs != null) getAllSubWebs();
            }
        }

        public void getFoldersSite(string nomSite)
        {
            //Met a jour le site choisis
            setWebWithString(nomSite);

            //Get the all list collection 
            ListCollection listColl = clientContext.Web.Lists;

            // Execute query. 
            clientContext.Load(listColl, lists => lists.Include(testList => testList.Title,
                                                                testList => testList.BaseTemplate));
            clientContext.ExecuteQuery();

            foreach (List list in listColl)
            {
                if (list.BaseTemplate == 101) // id dossier
                {
                    listDesDossier.Add(list);
                }

            }

        }

        public void getFolderContentTypes(string nomListe)
        {


            // Get the content type collection for the list nomListe
            nomDossier = nomListe;
            ContentTypeCollection contentTypeColl = clientContext.Web.Lists.GetByTitle(nomDossier).ContentTypes;
            //Execute the reques
            clientContext.Load(contentTypeColl);
            clientContext.ExecuteQuery();

            foreach (ContentType c in contentTypeColl)
            {
                listDesContentType.Add(c);
            }
        }

        public void getChampsDunContentType(string nomContentType)
        {
            nomChamp = nomContentType;
            //// Get the field  collection for the content type nomContentType contenu dans la collection contentTypeCollection
            foreach (ContentType ct in listDesContentType)
            {
                if (ct.Name == nomChamp)
                {
                    //Recupération des champs 
                    FieldCollection fieldColl = ct.Fields;
                    //Execution de la requette
                    clientContext.Load(fieldColl);
                    clientContext.ExecuteQuery();

                    foreach (Field f in fieldColl)
                    {
                        if (f.FromBaseType == false) listDesField.Add(f);
                    }
                }

            }

        }

        //Pour convertir la collection il faut mettre null dans les autres paramètres 

        public void setCollValue(string nomColl, object valeur)
        {
            nomChamp = nomColl;
            valeurChamp = valeur;
            fichier[nomChamp] = valeurChamp;
            fichier.Update(); // important, rembeber changes

            clientContext.ExecuteQuery();

        }

        public void uploadFile()
        {

            // Add the ListItem
            Folder folder = clientContext.Web.GetFolderByServerRelativeUrl(clientContext.Url + nomDossier);
            FileCreationInformation fci = new FileCreationInformation();
            fci.Content = System.IO.File.ReadAllBytes(filePath);
            string[] cut = filePath.Split('/');
            fci.Url = cut.Last();
            fci.Overwrite = true;

            File fileToUpload = folder.Files.Add(fci);
            clientContext.Load(fileToUpload);

            fichier = fileToUpload.ListItemAllFields;
            clientContext.Load(fichier);
            setCollValue("ContentTypeId", typeDuFichier.Id);
            string[] titre = cut.Last().Split('.');
            setCollValue("Title", titre.First());
            // Now invoke the server, just one time

            clientContext.ExecuteQuery();

        }

        public void setContentTypeWithString(string contentString)
        {
            foreach (ContentType contentType in listDesContentType)
            {
                if (contentType.Name == contentString) typeDuFichier = contentType;
            }

        }

        public void setWebWithString(string contentString, List<Web> webColl)
        {
            foreach (Web web in webColl)
            {
                if (web.Title == contentString) site = web;
            }

        }


        public void setWebWithString(string webString)
        {
            foreach (Web web in listDesSites)
            {
                if (web.Title == webString) site = web;
            }

            //Update the client context with the selected site
            clientContext = new ClientContext(domaine + site.ServerRelativeUrl);

        }
        public static List<string> convertToString(object collection)
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
                    listARetourner.Add(field.Title);
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

            return listARetourner;
        }
    }
}
