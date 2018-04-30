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
            //// Get the content type collection for the list "Custom"
            ContentTypeCollection contentTypeColl = clientContext.Web.Lists.GetByTitle(nomListe).ContentTypes;

            clientContext.Load(contentTypeColl);
            clientContext.ExecuteQuery();

            return contentTypeColl;
        }

        public static FieldCollection recupChampDunContentType(ClientContext clientContext, ContentTypeCollection contentTypeColl, string nomContentType)
        {
            FieldCollection fC=null;
            //// Get the content type collection for the list "Custom"
            foreach (ContentType ct in contentTypeColl)
            {
                if(ct.Name==nomContentType)
                {
                fC = ct.Fields;
                clientContext.Load(fC);
                clientContext.ExecuteQuery();
                }

            }
            return fC;
        }
    }
}
