using Microsoft.AspNetCore.Components.WebAssembly.Authentication;
using Microsoft.Authentication.WebAssembly.Msal.Models;
using Microsoft.Graph;
using System.Net.Http.Headers;

namespace CPR.Blazor.Wasm.Settings
{
    public static class MsGraphClientSetting
    {
        public static IServiceCollection AddGraphClient(this IServiceCollection services, params string[] scopes)
        {
            services.Configure<RemoteAuthenticationOptions<MsalProviderOptions>>(
                options =>
                {
                    foreach (var scope in scopes)
                    {
                        options.ProviderOptions.AdditionalScopesToConsent.Add(scope);
                    }
                });

            services.AddScoped<IAuthenticationProvider,
                NoOpGraphAuthenticationProvider>();
            services.AddScoped<IHttpProvider, HttpClientHttpProvider>(sp =>
                new HttpClientHttpProvider(new HttpClient()));
            services.AddScoped(sp =>
            {
                return new GraphServiceClient(
                    sp.GetRequiredService<IAuthenticationProvider>(),
                    sp.GetRequiredService<IHttpProvider>());
            });

            return services;
        }

        private class NoOpGraphAuthenticationProvider : IAuthenticationProvider
        {
            private readonly string graphUrl = "https://graph.microsoft.com/";

            public NoOpGraphAuthenticationProvider(IAccessTokenProvider tokenProvider)
            {
                TokenProvider = tokenProvider;
            }

            public IAccessTokenProvider TokenProvider { get; }

            public async Task AuthenticateRequestAsync(HttpRequestMessage request)
            {
                var result = await TokenProvider.RequestAccessToken(
                    new AccessTokenRequestOptions()
                    {
                        Scopes = new[] {
                            graphUrl + "User.Read",
                            graphUrl + "Files.Read",
                            graphUrl + "Files.Read.All",
                            graphUrl + "Files.ReadWrite",
                            graphUrl + "Files.ReadWrite.All"
                        }
                    });

                if (result.TryGetToken(out var token))
                {
                    request.Headers.Authorization ??= new AuthenticationHeaderValue(
                        "Bearer", token.Value);
                }
            }
        }

        private class HttpClientHttpProvider : IHttpProvider
        {
            private readonly HttpClient http;

            public HttpClientHttpProvider(HttpClient http)
            {
                this.http = http;
            }

            public ISerializer Serializer { get; } = new Serializer();

            public TimeSpan OverallTimeout { get; set; } = TimeSpan.FromSeconds(300);

            public void Dispose()
            {
            }

            public Task<HttpResponseMessage> SendAsync(HttpRequestMessage request)
            {
                return http.SendAsync(request);
            }

            public Task<HttpResponseMessage> SendAsync(HttpRequestMessage request,
                HttpCompletionOption completionOption,
                CancellationToken cancellationToken)
            {
                return http.SendAsync(request, completionOption, cancellationToken);
            }
        }
    }
}
