using CellServe.ExcelHandler;
using CellServe.ExcelHandler.Interfaces;
using CellServe.ExcelHandler.Strategies;

var builder = WebApplication.CreateBuilder(args);
var allowedOrigins = builder.Configuration.GetSection("Cors:AllowedOrigins").Get<string[]>() ?? [];

builder.Services.Configure<WorkbookOptions>(builder.Configuration.GetSection("Workbook"));
builder.Services.AddMemoryCache();
builder.Services.AddScoped<IWorkbookRepository, WorkbookRepository>();
builder.Services.AddScoped<ISheetFilterStrategy, SheetFilterStrategy>();
builder.Services.AddScoped<IRowModelingStrategy, RowModelingStrategy>();
builder.Services.AddControllersWithViews().AddRazorRuntimeCompilation();
if (allowedOrigins.Length > 0)
{
    builder.Services.AddCors(options => options.AddPolicy("ConfiguredOrigins", policy =>
        policy.WithOrigins(allowedOrigins).AllowAnyHeader().AllowAnyMethod()));
}

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseRouting();

if (allowedOrigins.Length > 0)
{
    app.UseCors("ConfiguredOrigins");
}

app.UseAuthorization();

app.UseStaticFiles();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");


app.Run();
