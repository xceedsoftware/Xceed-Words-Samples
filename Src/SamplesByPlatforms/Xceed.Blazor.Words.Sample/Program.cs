using Xceed.Blazor.Words.Sample.Components;
using Xceed.Blazor.Words.Sample.Services;

// Replace the License Key by a valid license.
Xceed.Words.NET.Licenser.LicenseKey = "XXXXX-XXXXX-XXXXX-YYYY";

var builder = WebApplication.CreateBuilder( args );

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.AddScoped<WordCreator>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if( !app.Environment.IsDevelopment() )
{
  app.UseExceptionHandler( "/Error", createScopeForErrors: true );
}

app.UseStaticFiles();
app.UseAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
