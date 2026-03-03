using Npgsql;
using System;
using System.Data;

var builder = WebApplication.CreateBuilder(args);

// Load config từ appsettings.json + appsettings.{ENV}.json
builder.Configuration
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
    .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true, reloadOnChange: true)
    .AddEnvironmentVariables(); // Ưu tiên biến môi trường (Render sẽ inject)

// Lấy PORT từ Render, nếu không có thì dùng 5000
var port = Environment.GetEnvironmentVariable("PORT") ?? "5000";
builder.WebHost.UseUrls($"http://0.0.0.0:{port}");


// Load cấu hình CORS từ appsettings
var allowedOrigins = builder.Configuration
    .GetSection("Cors:AllowedOrigins")
    .Get<string[]>();

builder.Services.AddCors(options =>
{
    options.AddPolicy("AppCorsPolicy", policy =>
    {
        policy.WithOrigins(allowedOrigins!)
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

// Lấy connection string từ biến môi trường (Render inject vào Configuration)
var connectionString = builder.Configuration.GetConnectionString("DefaultConnection");

if (string.IsNullOrEmpty(connectionString))
{
    throw new InvalidOperationException("⚠️ Connection string không được cấu hình. Kiểm tra biến môi trường 'ConnectionStrings__DefaultConnection'.");
}

builder.Services.AddScoped<IDbConnection>(sp =>
    new NpgsqlConnection(connectionString)); // Dùng PostgreSQL

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();

var app = builder.Build();

// (Tùy chọn) Tự động chuyển hướng HTTP → HTTPS
if (!app.Environment.IsDevelopment())
{
    app.UseHttpsRedirection();
}

app.UseCors("AppCorsPolicy");

app.UseAuthorization();
app.MapControllers();
app.Run();
