using System;
using System.IO;
using System.Net;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Afonsoft.Web.Library
{
    /// <summary>
    /// Page Base para as aplicações Web
    /// </summary>
    public class PageBase : Page
    {
        /// <summary>
        /// Recupere o objeto CultureInfo do Brasil, para usar no Convert.ToDateTime
        /// </summary>
        public System.Globalization.CultureInfo CultureInfo { get; } = new System.Globalization.CultureInfo("pt-BR");


        #region Compressor ViewState
        private readonly ObjectStateFormatter _objectStateFormatter = new ObjectStateFormatter();

        /// <summary>
        /// SavePageStateToPersistenceMedium - Compress Gzip ViewState
        /// </summary>
        protected override void SavePageStateToPersistenceMedium(object viewState)
        {
            try
            {
                byte[] viewStateArray;
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    _objectStateFormatter.Serialize(memoryStream, viewState);
                    viewStateArray = memoryStream.ToArray();
                }
                base.SavePageStateToPersistenceMedium(Convert.ToBase64String(Compressor.Compress(viewStateArray)));
            }
            catch
            {
                base.SavePageStateToPersistenceMedium(viewState);
            }
        }

        /// <summary>
        /// LoadPageStateFromPersistenceMedium - Decompress Gzip ViewState
        /// </summary>
        protected override object LoadPageStateFromPersistenceMedium()
        {
            try
            {
                string viewState = base.LoadPageStateFromPersistenceMedium().ToString() != "System.Web.UI.Pair" ? base.LoadPageStateFromPersistenceMedium().ToString() : ((Pair)base.LoadPageStateFromPersistenceMedium()).Second.ToString();
                byte[] bytes = Convert.FromBase64String(viewState);
                bytes = Compressor.Decompress(bytes);
                return _objectStateFormatter.Deserialize(Convert.ToBase64String(bytes));
            }
            catch
            {
                return base.LoadPageStateFromPersistenceMedium();
            }
        }
        #endregion

        #region IsAjax
        /// <summary>
        /// IsAjax - Verifica se a requisição vem do Ajax
        /// </summary>
        public bool IsAjax
        {
            get
            {
                if (Request.Headers["HTTP_X_MICROSOFTAJAX"] != null)
                    return true;
                if (Request.Headers["X-MicrosoftAjax"] == "Delta=true")
                    return true;
                return false;

            }
        }
        #endregion

        #region Remote IP/Logon/Page

        /// <summary>
        /// Recuperar o Ip do client conectado no portal 
        /// </summary>
        /// <returns>IP</returns>
        public string GetIPAddress()
        {
            return RemoteAddressIp;
        }

        /// <summary>
        /// Recuperar o Nome da Pagina que está sendo acessada
        /// </summary>
        /// <returns>xxx.aspx</returns>
        public string GetPage()
        {
            return Path.GetFileName(HttpContext.Current.Request.Url.AbsolutePath);
        }

        /// <summary>
        /// Recueprar o IP da maquina que está fazendo a requisição
        /// </summary>
        public string RemoteAddressIp
        {
            get
            {
                string ip = "";
                try
                {
                    ip = HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];

                    if (string.IsNullOrEmpty(ip) || ip == "::1" || ip.StartsWith("127"))
                        ip = HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];

                    if (string.IsNullOrEmpty(ip) || ip == "::1" || ip.StartsWith("127"))
                        ip = HttpContext.Current.Request.ServerVariables["REQUEST_ADDR"];

                    if (string.IsNullOrEmpty(ip) || ip == "::1" || ip.StartsWith("127"))
                        ip = HttpContext.Current.Request.ServerVariables["REMOTE_HOST"];

                    if (string.IsNullOrEmpty(ip) || ip == "::1" || ip.StartsWith("127"))
                        ip = HttpContext.Current.Request.UserHostAddress;
                }
                catch
                {
                    // ignored
                }

                if (!string.IsNullOrEmpty(ip))
                {
                    string[] addresses = ip.Split(',');
                    if (addresses.Length != 0)
                    {
                        return addresses[0];
                    }
                }

                return ip;
            }
        }

        /// <summary>
        /// Recuperar o nome da Maquina que está fazendo a requisição
        /// </summary>
        public string RemoteHost
        {
            get
            {
                string hostName = "";
                string ip = RemoteAddressIp;
                try
                {
                    if (!string.IsNullOrEmpty(ip))
                    {
                        hostName = Dns.GetHostEntry(ip).HostName;
                        if (string.IsNullOrEmpty(hostName) || hostName == ip)
                            hostName = HttpContext.Current.Request.ServerVariables["HTTP_HOST"];
                    }
                    else
                    {
                        hostName = HttpContext.Current.Request.ServerVariables["HTTP_HOST"];
                    }
                }
                catch
                {
                    // ignored
                }
                return hostName;
            }
        }

        /// <summary>
        /// Recuperar o Logon da Maquina que está fazendo a requisição
        /// </summary>
        public string RemoteLogon
        {
            get
            {
                string logon = "";
                try
                {
                    logon = Request.ServerVariables["LOGON_USER"];

                    if (string.IsNullOrEmpty(logon) || logon.IndexOf("userportal", StringComparison.Ordinal) > 0)
                        logon = Request.ServerVariables["AUTH_USER"];

                    if (string.IsNullOrEmpty(logon) || logon.IndexOf("userportal", StringComparison.Ordinal) > 0)
                        logon = Request.ServerVariables["REMOTE_USER"];

                    if (string.IsNullOrEmpty(logon) || logon.IndexOf("userportal", StringComparison.Ordinal) > 0)
                        logon = Request.LogonUserIdentity?.Name;

                    if (string.IsNullOrEmpty(logon) || logon.IndexOf("userportal", StringComparison.Ordinal) > 0)
                        logon = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                }
                catch
                {
                    // ignored
                }
                return logon;
            }
        }
        #endregion

        #region GetPostBack
        /// <summary>
        /// Recuperar o nome do Control que causou o PostBack
        /// </summary>
        public string GetPostBackControl()
        {
            string outupt = "";

            //get the __EVENTTARGET of the Control that cased a PostBack(except Buttons and ImageButtons)
            string targetCtrl = Page.Request.Params.Get("__EVENTTARGET");
            if (!string.IsNullOrEmpty(targetCtrl))
            {
                outupt = targetCtrl;
            }
            else
            {
                //if button is cased a postback
                foreach (string str in Request.Form)
                {
                    Control ctrl = Page.FindControl(str);
                    if (!(ctrl is Button)) continue;
                    outupt = ctrl.ID;
                    break;
                }
            }
            return outupt;
        }

        /// <summary>
        /// Recuperar o Argumento do PostBack
        /// </summary>
        public string GetPostBackArgument()
        {
            string outupt = "";

            //get the __EVENTTARGET of the Control that cased a PostBack(except Buttons and ImageButtons)
            string targetCtrl = Page.Request.Params.Get("__EVENTARGUMENT");
            if (!string.IsNullOrEmpty(targetCtrl))
            {
                outupt = targetCtrl;
            }
            else
            {
                //if button is cased a postback
                foreach (string str in Request.Form)
                {
                    Control ctrl = Page.FindControl(str);
                    if (!(ctrl is Button)) continue;
                    outupt = ((Button)ctrl).CommandArgument;
                    break;
                }
            }
            return outupt;
        }
        #endregion

        #region Alert
        /// <summary>
        /// Exibi um alert na tela do sistema
        /// </summary>
        public void Alert(string text)
        {
            Alert(text, null);
        }

        /// <summary>
        /// Exibi um alert na tela do sistema
        /// </summary>
        public void Alert(Exception exception)
        {
            Alert(exception.Message, exception);
        }

        /// <summary>
        /// Exibi um alert na tela do sistema
        /// </summary>
        public void Alert(string text, Exception exception)
        {
            Alert(text, exception, this);
        }
        /// <summary>
        /// Exibi um alert na tela do sistema
        /// </summary>
        public void Alert(string text, Exception exception, Control sender)
        {
            if (sender == null)
                sender = this;


            Exception tmpEx = exception;
            string msgExeption="";

            while (tmpEx != null)
            {
                msgExeption += tmpEx.Message
                                   .Replace("'", "`")
                                   .Replace(Environment.NewLine, "<br/>")
                                   .Replace("\n", "<br/>")
                                   .Replace("\r", "<br/>")
                                   .Replace("\\", "/") + "<br/>";
                tmpEx = tmpEx.InnerException;
            }

            string alert = text + Environment.NewLine + msgExeption;
            alert = alert.Replace("'", "`")
                .Replace(Environment.NewLine, "<br/>")
                .Replace("\n", "<br/>")
                .Replace("\r", "<br/>")
                .Replace("\\", "/")
                .Replace("'", "`")
                .Trim();

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("try { ");
            sb.AppendLine("   jQuery(document).ready(function() {");
            sb.AppendLine("     try { ");
            sb.AppendFormat("       bootbox.alert({title: \"{0}\",  message: \"{1}\"});", "Alerta",alert);
            sb.AppendLine("");
            sb.AppendLine("     }catch(e){ ");
            sb.AppendFormat("       alert('{0}');", alert);
            sb.AppendLine("");
            sb.AppendLine("     }");
            sb.AppendLine("   });");
            sb.AppendLine("}catch(e){ ");
            sb.AppendFormat("    alert('{0}');", alert);
            sb.AppendLine("");
            sb.AppendLine("}");

            ScriptManager.RegisterStartupScript(sender, sender.GetType(), "MsgAlert", sb.ToString(), true);
        }

        #endregion
    }
}
