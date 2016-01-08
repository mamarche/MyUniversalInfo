using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;
using Windows.Security.Authentication.Web.Core;
using Windows.Security.Credentials;
using Windows.Storage;
using Windows.UI.Popups;

namespace MyUniversalInfo
{
    internal static class OfficeHelper
    {
        static string serviceEndpoint = "https://graph.microsoft.com/v1.0/";
        static string clientId = App.Current.Resources["ida:ClientId"].ToString();

        //autority da utilizzare se l'app lavora su uno specifico Tenant
        static string authority = string.Format("{0}{1}", App.Current.Resources["ida:AADInstance"].ToString(),
                                                          App.Current.Resources["ida:Domain"].ToString());

        //autority da utilizzare se l'app lavora su ogni Azure Tenant
        //static string authority = "organizations";

        static string ResourceUrl = "https://graph.microsoft.com/";

        /// <summary>
        /// Restituisce l'access Token per le Graph APIs
        /// </summary>
        public static async Task<string> GetTokenAsync()
        {
            WebAccount userAccount = null;

            WebAccountProvider aadAccountProvider = await WebAuthenticationCoreManager.FindAccountProviderAsync("https://login.microsoft.com", authority);

            WebTokenRequest webTokenRequest = new WebTokenRequest(aadAccountProvider, string.Empty, clientId, WebTokenRequestPromptType.ForceAuthentication);
            webTokenRequest.Properties.Add("resource", ResourceUrl);

            WebTokenRequestResult webTokenRequestResult = await WebAuthenticationCoreManager.RequestTokenAsync(webTokenRequest);
            if (webTokenRequestResult.ResponseStatus == WebTokenRequestStatus.Success)
            {
                WebTokenResponse webTokenResponse = webTokenRequestResult.ResponseData[0];
                userAccount = webTokenResponse.WebAccount;
                return webTokenResponse.Token;
            }

            return null;
        }

        /// <summary>
        /// Restituisce i dati dell'utente
        /// </summary>
        public static async Task<Microsoft.Graph.User> GetMyInfoAsync()
        {
            //JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await OfficeHelper.GetTokenAsync();
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {token}");

                // Endpoint dell'utente loggato
                Uri usersEndpoint = new Uri($"{serviceEndpoint}me");

                HttpResponseMessage response = await client.GetAsync(usersEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    return Newtonsoft.Json.JsonConvert.DeserializeObject<Microsoft.Graph.User>(responseContent);                    
                }
                else
                {
                    var msg = new MessageDialog("Non è stato possibile recuperare i dati dell'utente. Risposta del server: " + response.StatusCode);
                    await msg.ShowAsync();
                    return null;
                }
            }
            catch (Exception e)
            {
                var msg = new MessageDialog("Si è verificato il seguente errore: " + e.Message);
                await msg.ShowAsync();
                return null;
            }
        }

        /// <summary>
        /// Restituisce la Redirect Uri dell'applicazione da impostare sul portale di Azure
        /// </summary>
        public static string GetRedirectUri()
        {
            return string.Format("ms-appx-web://microsoft.aad.brokerplugin/{0}", 
                WebAuthenticationBroker.GetCurrentApplicationCallbackUri().Host).ToUpper();
        }

    }
}
