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
        private string Domaine { get; set; }
        private string FilePath { get; set; }
        public string NomDossier { get; set; }
        public string NomChamp { get; set; }
        public object ValeurChamp { get; set; }

        private Web Site { get; set; }
        private ContentType TypeDuFichier { get; set; }
        private ClientContext ClientCtx { get; set; }
        private ListItem Fichier { get; set; }


        public List<Web> ListDesSites { get; set; }
        public List<List> ListDesDossier { get; set; }
        public List<ContentType> ListDesContentType { get; set; }
        public List<Field> ListDesField { get; set; }

        /* Peux être inutile
        public List<string> ListDesSitesString { get; set; }
        List<string> ListDesDossierString;
        List<string> contentTypeCollString;
        List<string> fieldCollString;
        */

        public InstanceBot(string url, string chemin)
        {
            ClientCtx = new ClientContext(url);
            FilePath = chemin;
            string[] split = ClientCtx.Url.Split('/');
            Domaine = split[0] + "//" + split[2];

            ListDesSites = new List<Web>();
            ListDesDossier = new List<List>();
            ListDesContentType = new List<ContentType>();
            ListDesField = new List<Field>();

            GetAllSubWebs();

        }
        public void GetAllSubWebs()
        {
            // Get the SharePoint web  
            Web web = ClientCtx.Web;
            ClientCtx.Load(web, website => website.Webs, website => website.Title);

            // Execute the query to the server  
            ClientCtx.ExecuteQuery();

            // Loop through all the webs  
            foreach (Web subWeb in web.Webs)
            {
                string newpath = Domaine + subWeb.ServerRelativeUrl;
                ListDesSites.Add(subWeb);
                ClientCtx = new ClientContext(newpath);
                if (subWeb.Webs != null) GetAllSubWebs();
            }
        }

        public void GetFoldersSite(string nomSite)
        {
            //Met a jour le site choisis
            SetWebWithString(nomSite);

            //Get the all list collection 
            ListCollection listColl = ClientCtx.Web.Lists;

            // Execute query. 
            ClientCtx.Load(listColl, lists => lists.Include(testList => testList.Title,
                                                                testList => testList.BaseTemplate));
            ClientCtx.ExecuteQuery();

            foreach (List list in listColl)
            {
                if (list.BaseTemplate == 101) // id dossier
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
            ClientCtx.ExecuteQuery();

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
                if (ct.Name == NomChamp)
                {
                    //Recupération des champs 
                    FieldCollection fieldColl = ct.Fields;
                    //Execution de la requette
                    ClientCtx.Load(fieldColl);
                    ClientCtx.ExecuteQuery();

                    foreach (Field f in fieldColl)
                    {
                        if (f.FromBaseType == false) ListDesField.Add(f);
                    }
                }

            }

        }

        //Pour convertir la collection il faut mettre null dans les autres paramètres 

        public void SetCollValue(string nomColl, object valeur)
        {
            NomChamp = nomColl;
            ValeurChamp = valeur;
            Fichier[NomChamp] = ValeurChamp;
            Fichier.Update(); // important, rembeber changes

            ClientCtx.ExecuteQuery();

        }

        public void ToUploadFile()
        {

            // Add the ListItem
            if (NomDossier == "Documents") NomDossier = "Shared Documents";
            Folder folder = ClientCtx.Web.GetFolderByServerRelativeUrl(ClientCtx.Url + "/" + NomDossier);
            FileCreationInformation fci = new FileCreationInformation();
            fci.Content = System.IO.File.ReadAllBytes(FilePath);
            string[] cut = FilePath.Split('/');
            fci.Url = cut.Last();
            fci.Overwrite = true;

            File fileToUpload = folder.Files.Add(fci);
            ClientCtx.Load(fileToUpload);

            Fichier = fileToUpload.ListItemAllFields;
            ClientCtx.Load(Fichier);
            SetCollValue("ContentTypeId", TypeDuFichier.Id);
            string[] titre = cut.Last().Split('.');
            SetCollValue("Title", titre.First());
            // Now invoke the server, just one time

            ClientCtx.ExecuteQuery();
        }

        public void SetContentTypeWithString(string contentString)
        {
            foreach (ContentType contentType in ListDesContentType)
            {
                if (contentType.Name == contentString) TypeDuFichier = contentType;
            }

        }

        public void SetWebWithString(string contentString, List<Web> webColl)
        {
            foreach (Web web in webColl)
            {
                if (web.Title == contentString) Site = web;
            }

        }


        public void SetWebWithString(string webString)
        {
            foreach (Web web in ListDesSites)
            {
                if (web.Title == webString) Site = web;
            }

            //Update the client context with the selected site
            ClientCtx = new ClientContext(Domaine + Site.ServerRelativeUrl);

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
