using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Microsoft.JSInterop;
using System.Globalization;

namespace CPR.Blazor.Wasm.Settings
{
    public static class ConfigureSetting
    {
        public static async Task LoadSettingJson(this WebAssemblyHostConfiguration configuration, HttpClient httpClient)
        {
            using var response = await httpClient.GetAsync("yusettings.json");
            using var stream = await response.Content.ReadAsStreamAsync();

            configuration.AddJsonStream(stream);
        }

        public static void SetMsAuth(this WebAssemblyHostBuilder builder)
        {
            builder.Services.AddMsalAuthentication(options =>
            {
                options.ProviderOptions.Authentication.ClientId = builder.Configuration["MsAuth:ClientID"];
                //options.ProviderOptions.Authentication.Authority =
                //    string.Format("https://login.microsoftonline.com/{0}", Constant.MS_AUTH_TENANT_ID);
                //Microsoft identity platform v2.0才支援多租用戶，但目前不支援&prompt=select_account選擇帳號
                //Azure AD 資訊清單 - allowPublicClient:true, signInAudience: AzureADandPersonalMicrosoftAccount
                options.ProviderOptions.Authentication.Authority = "https://login.microsoftonline.com/common";
                options.ProviderOptions.Authentication.ValidateAuthority = true;
                options.ProviderOptions.LoginMode = "redirect";
                options.ProviderOptions.DefaultAccessTokenScopes.Add("https://graph.microsoft.com/User.Read");
                options.ProviderOptions.DefaultAccessTokenScopes.Add("https://graph.microsoft.com/Files.Read");
                options.ProviderOptions.DefaultAccessTokenScopes.Add("https://graph.microsoft.com/Files.Read.All");
                options.ProviderOptions.DefaultAccessTokenScopes.Add("https://graph.microsoft.com/Files.ReadWrite");
                options.ProviderOptions.DefaultAccessTokenScopes.Add("https://graph.microsoft.com/Files.ReadWrite.All");
            });
        }

        public static async Task SetCulture(this WebAssemblyHost host)
        {
            CultureInfo culture;
            var js = host.Services.GetRequiredService<IJSRuntime>();
            var result = await js.InvokeAsync<string>("blazorCulture.get");

            culture = new CultureInfo("zh-Hant");
            await js.InvokeVoidAsync("blazorCulture.set", "zh-Hant");
            /*if (result != null)
            {
                culture = new CultureInfo(result);
            }
            else
            {
                culture = new CultureInfo("zh-Hant");
                await js.InvokeVoidAsync("blazorCulture.set", "zh-Hant");
            }*/

            CultureInfo.DefaultThreadCurrentCulture = culture;
            CultureInfo.DefaultThreadCurrentUICulture = culture;
        }
    }
}
