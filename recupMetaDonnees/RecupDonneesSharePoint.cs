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
        public static List<string> convertToString(ListCollection listColl = null, ContentTypeCollection contentTypeColl = null, FieldCollection fieldColl = null)
        {
            List<string> listARetourner = new List<string>();

            if (listColl!=null)
            {
                foreach(List list in listColl)
                {
                    listARetourner.Add(list.Title);
                }
            }
            else if (contentTypeColl != null)
            {
                foreach (ContentType contentType in contentTypeColl)
                {
                    listARetourner.Add(contentType.Name);
                }
            }
            else if (fieldColl != null)
            {
                foreach (Field field in fieldColl)
                {
                    listARetourner.Add(field.Title);
                }
            }

            return listARetourner;
        }
    }
}
