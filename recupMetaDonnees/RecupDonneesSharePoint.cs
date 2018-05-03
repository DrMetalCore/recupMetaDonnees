using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace recupMetaDonnees
{
    class RecupDonneesSharePoint
    {
        public static ListCollection getAllListes(ClientContext clientContext)
        {
            //Get the all list collection 
            ListCollection listColl = clientContext.Web.Lists;

            // Execute query. 
            clientContext.Load(listColl, lists => lists.Include(testList => testList.Title));
            clientContext.ExecuteQuery();

            return listColl;
        }

        public static ContentTypeCollection getContentTypesDuneList(ClientContext clientContext, string nomListe)
        {
            //// Get the content type collection for the list nomListe
            ContentTypeCollection contentTypeColl = clientContext.Web.Lists.GetByTitle(nomListe).ContentTypes;

            //Execute the reques
            clientContext.Load(contentTypeColl);
            clientContext.ExecuteQuery();

            return contentTypeColl;
        }

        public static FieldCollection getChampsDunContentType(ClientContext clientContext, ContentTypeCollection contentTypeColl, string nomContentType)
        {
            //Initialisaton des variable necessaires 
            FieldCollection fC=null;
            //// Get the field  collection for the content type nomContentType contenu dans la collection contentTypeCollection
            foreach (ContentType ct in contentTypeColl)
            {
                if(ct.Name==nomContentType)
                {
                    //Recupération des champs 
                    fC = ct.Fields;
                    //Execution de la requette
                    clientContext.Load(fC);
                    clientContext.ExecuteQuery();
                }

            }
            return fC;
        }
        
        //Pour convertir la collection il faut mettre null dans les autres paramètres 
        public static List<string> convertToString(object collection)
        {
            List<string> listARetourner = new List<string>();
            
            if (collection.GetType().ToString()== "Microsoft.SharePoint.Client.ListCollection")
            {
                ListCollection collectionConverti = (ListCollection)collection;
                foreach (List list in collectionConverti)
                {
                    listARetourner.Add(list.Title);
                }
            }
            else if (collection.GetType().ToString() == "Microsoft.SharePoint.Client.ContentTypeCollection")
            {
                ContentTypeCollection collectionConverti = (ContentTypeCollection)collection;
                foreach (ContentType contentType in collectionConverti)
                {
                    listARetourner.Add(contentType.Name);
                }
            }
            else if (collection.GetType().ToString() == "Microsoft.SharePoint.Client.FieldCollection")
            {
                FieldCollection collectionConverti = (FieldCollection)collection;
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
        
        public static void setCollValue(ClientContext clientContext, ListItem item, string nomColl, object valeur)
        {

            item[nomColl] = valeur;
            item.Update(); // important, rembeber changes

            clientContext.ExecuteQuery();
        }

        public static ListItem uploadFile(ClientContext clientContext, string filePath, string nomDossier, ContentType contentType)
        {
            // Add the ListItem
            Folder folder = clientContext.Web.GetFolderByServerRelativeUrl(clientContext.Url + nomDossier);
            FileCreationInformation fci = new FileCreationInformation();
            fci.Content = System.IO.File.ReadAllBytes("../../" + filePath);
            fci.Url = filePath;
            fci.Overwrite = true;
                    
            File fileToUpload = folder.Files.Add(fci);
            clientContext.Load(fileToUpload);

            ListItem item = fileToUpload.ListItemAllFields;
            clientContext.Load(item);

            setCollValue(clientContext, item, "ContentTypeId", contentType.Id);

            // Now invoke the server, just one time
            clientContext.ExecuteQuery();

            return item;
            
        }

        public static ContentType StringToContentType(string contentString, ContentTypeCollection contentTypesColl)
        {
            ContentType contentTypeARetourner = null;
            foreach (ContentType contentType in contentTypesColl)
            {
                if (contentType.Name == contentString) contentTypeARetourner = contentType;
            }
            return contentTypeARetourner;
        }

        public static Web StringToWeb(string contentString, List<Web> webColl)
        {
            Web webARetourner = null;
            foreach (Web web in webColl)
            {
                if (web.Title == contentString) webARetourner = web;
            }
            return webARetourner;
        }

        public static List<Web> GetAllSubWebs(ClientContext clientContext)
        {
            List<Web> listARetourner = new List<Web>();
            
            // Get the SharePoint web  
            Web web = clientContext.Web;
            clientContext.Load(web, website => website.Webs, website => website.Title);
            
            // Execute the query to the server  
            clientContext.ExecuteQuery();
            string[] split = clientContext.Url.Split('/');
            string domain = split[0] + "//" + split[2];

            // Loop through all the webs  
            foreach (Web subWeb in web.Webs)
            {
                string newpath = domain + subWeb.ServerRelativeUrl;
                listARetourner.Add(subWeb);
                clientContext = new ClientContext(newpath);
                listARetourner = listARetourner.Concat(GetAllSubWebs(clientContext)).ToList();
            }
            return listARetourner;
            
        }
    }
}
