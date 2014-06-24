using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aspose.EmailProcessing.Library;

namespace Aspose.EmailProcessing
{
    public partial class Default1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (HttpContext.Current.Session[Constants.MailHelperSession] != null)
            {
                HttpContext.Current.Session.Remove(Constants.MailHelperSession);
            }
        }
    }
}