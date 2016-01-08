using Microsoft.Office365.OutlookServices;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Security.Authentication.Web.Core;
using Windows.Security.Credentials;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace MyUniversalInfo
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {

        public MainPage()
        {
            this.InitializeComponent();

            //visualizzo l'uri da impostare su Azure per il redirect
            redirectUriText.Text = OfficeHelper.GetRedirectUri();
        }

        private async void getInfoButton_Click(object sender, RoutedEventArgs e)
        {
            var userData = await OfficeHelper.GetMyInfoAsync();

            givennameText.Text = $"Nome: {userData.givenName}";
            surnameText.Text = $"Cognome {userData.surname}";
            emailText.Text = $"E-mail: {userData.mail}";
        }
    }
}
