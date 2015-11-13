using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertToPdf.Common
{
    public static class CommonUtilities
    {
        public static SPList isValidLibrary(string libraryName)
        {
            SPList library = null;
            var web = SPContext.Current.Web;
            if (String.IsNullOrEmpty(libraryName))
            {
                return library;
            }

            var list = web.Lists.TryGetList(libraryName);

            return list;
        }
    }
}
