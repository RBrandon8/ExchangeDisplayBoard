using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.Exchange.WebServices.Data;
using System.ComponentModel;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Security.Cryptography;
using System.Web.UI.HtmlControls;
using System.Drawing;

namespace ExchangeDisplayBoard
{
    public partial class MainForm : System.Web.UI.Page
    {
        // EWS integration variables
        private static Uri ewsURI = new Uri("https://outlook.office365.com/EWS/Exchange.asmx"); //Set for O365 only tested on an on-premise server.
        private static readonly string serviceAccount = "DisplayBoardSA";
        private static byte[] toEncrypt = UnicodeEncoding.ASCII.GetBytes("yourpasswordhere"); //Must be 16 characters
        private static readonly string defaultRoomList = "yourRoomList@domain.com";

        //Formatting
        private static Color alternatingRowColor = Color.FromArgb(232, 232, 232);
        //Color gradient for the div line.
        private static string sPLineLeftColor = "#0072C6";
        private static string sPLineCenterColor = "#A9A9A9";
        private static string sPLineRightColor = "#6c6c6c";

        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                if (!IsPostBack)
                {
                    //To display an image uncomment lines
                    //Image1.Visible = true;
                    //Image1.ImageUrl = "yourimage.jpg";

                    //Set CSS elements from formatting variables
                    HtmlGenericControl sPLine = (HtmlGenericControl)Page.FindControl("sPLine");
                    sPLine.Style.Add("background",
                        String.Format("linear-gradient(to right, {0} 18.20%,{1} 62.07%,{2} 100%)",
                        sPLineLeftColor, sPLineCenterColor, sPLineRightColor));

                    //Encrypt password in memory for reuse when page needs to be refreshed.
                    EncryptInMemoryData(toEncrypt, MemoryProtectionScope.SameLogon);
                    DataLoad();
                    Session["cnt"] = 0;
                }
            }            
            catch (Exception ex)
            {
                string error = ex.Message;
            }


        }

        
        protected void gvApps_DataBound(object sender, EventArgs e)
        {
        }

        protected void gvApps_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            // Format rows with alternating color and remove borders
            Color backgroundColor = alternatingRowColor;
            foreach (TableCell tc in e.Row.Cells)
            {
                tc.BorderStyle = BorderStyle.None;
            }

            if(e.Row.RowIndex % 2 == 0)
            {
                e.Row.Cells[1].BackColor = backgroundColor;
                e.Row.Cells[2].BackColor = backgroundColor;
                e.Row.Cells[3].BackColor = backgroundColor;
            }
        }

        protected void DataLoad()
        {
            //Certificate Validation not implemented.  Best practice to check for a true return.    
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            EWS.displayApps = new BindingList<Appt>();

            //Retrieve uri parameters for which board to display
            string uri = HttpContext.Current.Request.Url.Query;
            string boardParam = HttpUtility.ParseQueryString(uri).Get("board");

            //Connect to EWS
            ExchangeService service = new ExchangeService();
            DecryptInMemoryData(toEncrypt, MemoryProtectionScope.SameLogon);
            service.Credentials = new WebCredentials(serviceAccount, UnicodeEncoding.ASCII.GetString(toEncrypt));
            EncryptInMemoryData(toEncrypt, MemoryProtectionScope.SameLogon);
            service.Url = ewsURI;

            // Initialize values for the start and end times, and the number of appointments to retrieve.
            lblDateHeader.Text = DateTime.Today.ToLongDateString();
            DateTime startDate = DateTime.Now;
            DateTime endDate = DateTime.Today.AddHours(23).AddMinutes(59);

            //Display proper logo depending on board displayed
            String[] rooms = null;
            EWS.GetExchangeRoomList(service);
            EmailAddress email = null;

            //Default Board if no parameter is sent.
            if (boardParam is null)
            {
                email = EWS.roomsList.Where(x => x.Address == defaultRoomList).First();
                if (email != null)
                    EWS.GetExchangeRooms(service, email);
            }
            else
            {
                email = EWS.roomsList.Where(x => x.Address == boardParam).First();
                if (email != null)
                    EWS.GetExchangeRooms(service, email);
            }

            if (EWS.rooms.Count > 0)
            {
                if (email != null)
                    lblRoomName.Text = email.Name;
                else
                    lblRoomName.Text = "";

                foreach (EmailAddress room in EWS.rooms)
                {
                    EWS.GetResourceCalendarItems(service, room.Address, startDate, endDate);
                }
            }


            EWS.SortAppsList();
            gvApps.DataSource = EWS.displayApps;
            /* Test table formatting without Exchange integration
             new BindingList<Appt>() { new Appt("1",DateTime.Now, DateTime.Now.AddHours(1), "Meeting"),
                                       new Appt("1",DateTime.Now, DateTime.Now.AddHours(1), "Meeting"),
                                       new Appt("1",DateTime.Now, DateTime.Now.AddHours(1), "Meeting"),
                                       new Appt("1",DateTime.Now, DateTime.Now.AddHours(1), "Meeting"),
                                       new Appt("1",DateTime.Now, DateTime.Now.AddHours(1), "Meeting"),
                                       new Appt("1",DateTime.Now, DateTime.Now.AddHours(1), "Meeting"),
                                       new Appt("1",DateTime.Now, DateTime.Now.AddHours(1), "Meeting"),
                                       new Appt("1",DateTime.Now, DateTime.Now.AddHours(1), "Meeting"),
                                       new Appt("1",DateTime.Now, DateTime.Now.AddHours(1), "Meeting"),
                                       new Appt("1",DateTime.Now, DateTime.Now.AddHours(1), "Meeting"),
                                       new Appt("1",DateTime.Now, DateTime.Now.AddHours(1), "Meeting"),
                                       new Appt("1",DateTime.Now, DateTime.Now.AddHours(1), "Meeting"),  };*/
            gvApps.DataBind();
        }

        protected static bool CertificateValidationCallBack(
        object sender,
        System.Security.Cryptography.X509Certificates.X509Certificate certificate,
        System.Security.Cryptography.X509Certificates.X509Chain chain,
        System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }
            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid.
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }
                // When processing reaches this line, the only errors in the certificate chain are
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange Server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }

        /* 
         *   Encryption methods taken from Microsoft Docs.  Used to store password in memory for
         *   refresh calls.
         *   https://docs.microsoft.com/en-us/dotnet/standard/security/how-to-use-data-protection
         * 
         */
        public static void EncryptInMemoryData(byte[] Buffer, MemoryProtectionScope Scope)
        {
            if (Buffer == null)
                throw new ArgumentNullException("Buffer");
            if (Buffer.Length <= 0)
                throw new ArgumentException("Buffer");


            // Encrypt the data in memory. The result is stored in the same same array as the original data.
            ProtectedMemory.Protect(Buffer, Scope);

        }

        public static void DecryptInMemoryData(byte[] Buffer, MemoryProtectionScope Scope)
        {
            if (Buffer == null)
                throw new ArgumentNullException("Buffer");
            if (Buffer.Length <= 0)
                throw new ArgumentException("Buffer");


            // Decrypt the data in memory. The result is stored in the same same array as the original data.
            ProtectedMemory.Unprotect(Buffer, Scope);

        }

        public static byte[] CreateRandomEntropy()
        {
            // Create a byte array to hold the random value.
            byte[] entropy = new byte[16];

            // Create a new instance of the RNGCryptoServiceProvider.
            // Fill the array with a random value.
            new RNGCryptoServiceProvider().GetBytes(entropy);

            // Return the array.
            return entropy;
        }

        protected void btn1_Click(object sender, EventArgs e)
        {

        }

        protected void btn1_Click1(object sender, EventArgs e)
        {
        }

        protected void btn1_Click2(object sender, EventArgs e)
        {
        }


        /*
         * On Timer tick scrolls the table of appointments slowly to the bottom via JavaScript 
         * and then returns to the top.  After 30 ticks refreshes meetings to remove old meetings
         * and roll over to a new day. 
         */

        protected void Timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                //Intialize session cnt if null
                if (Session["cnt"] == null)
                    Session["cnt"] = 0;

                //Increment cnt and save into sessionn 
                Session["cnt"] = (int)Session["cnt"] + 1;
                int cnt = (int)Session["cnt"];

                //Calls AutoScrollEdge method in JavaScript method to auto-scroll table
                Page.ClientScript.RegisterStartupScript(this.GetType(), "myScript", "AutoScrollEdge();", true);

                //Check for new events and remove events that are over after 30 ticks
                if (cnt > 30)
                {
                    DataLoad();
                    Session["cnt"] = 0;
                }
            }
            catch(Exception ex)
            { }
        }

        protected void timerReload_Tick(object sender, EventArgs e)
        {
        }
    }
}