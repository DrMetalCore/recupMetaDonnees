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
        public static ContentTypeCollection recupContentTypeDuneList(ClientContext clientContext, string nomListe)
        {
            //// Get the content type collection for the list nomListe
            ContentTypeCollection contentTypeColl = clientContext.Web.Lists.GetByTitle(nomListe).ContentTypes;

            //Execute the reques
            clientContext.Load(contentTypeColl);
            clientContext.ExecuteQuery();

            return contentTypeColl;
        }

        public static FieldCollection recupChampDunContentType(ClientContext clientContext, ContentTypeCollection contentTypeColl, string nomContentType)
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
    }
}
