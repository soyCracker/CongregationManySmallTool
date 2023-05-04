using CPR.Blazor.Wasm;
using CPR.Blazor.Wasm.Settings;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using MudBlazor.Services;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

HttpClient httpClient = new HttpClient()
{
    BaseAddress = new Uri(builder.HostEnvironment.BaseAddress)
};
builder.Services.AddScoped(sp => httpClient);
await builder.Configuration.LoadSettingJson(httpClient);

builder.Services.AddMudServices();
builder.Services.AddGraphClient();
builder.SetMsAuth();

//»y¨t
builder.Services.AddLocalization(option =>
{
    option.ResourcesPath = "Resources";
});

var host = builder.Build();
//»y¨t
await host.SetCulture();
await host.RunAsync();
