using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Xceed.Blazor.Words.Sample;
using Xceed.Blazor.Words.Sample.Services;

// Replace the License Key by a valid license, change the values.
Xceed.Words.NET.Licenser.LicenseKey = "XXXXX-XXXXX-XXXXX-YYYY";

var builder = WebAssemblyHostBuilder.CreateDefault( args );
builder.RootComponents.Add<App>( "#app" );
builder.RootComponents.Add<HeadOutlet>( "head::after" );

builder.Services.AddScoped( sp => new HttpClient { BaseAddress = new Uri( builder.HostEnvironment.BaseAddress ) } );
builder.Services.AddScoped<WordCreator>();
await builder.Build().RunAsync();
