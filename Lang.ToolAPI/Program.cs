using Lang.ToolAPI.Filter;
using Lang.ToolBiz;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace Lang.ToolAPI
{
    public class Program
    {
        public static void Main(string[] args)
        {
           
            var builder = WebApplication.CreateBuilder(args);

            // Add services to the container.

            // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();
            builder.Services.AddCors(options =>
            {
                options.AddPolicy("cors", builder =>
                {
                    builder.SetIsOriginAllowed(p => true) //允许指定域名访问
                    .AllowAnyMethod()
                    .AllowAnyHeader()
                    .AllowCredentials();//指定处理cookie
                });
            });

            builder.Services.Configure<ApiBehaviorOptions>((o) =>
            {
                o.SuppressModelStateInvalidFilter = true;
            });

            builder.Services.AddResponseCompression();

            builder.Services.AddControllers(a => { a.Filters.Add(typeof(GlobalExceptionsFilter)); })
               .AddNewtonsoftJson(options =>
               {
                   options.SerializerSettings.DateFormatString = "yyyy-MM-dd HH:mm:ss";
                   options.SerializerSettings.NullValueHandling = NullValueHandling.Ignore;
               });

            builder.Services.AddSingleton<BPDF>();

           var app = builder.Build();

            // Configure the HTTP request pipeline.
            if (app.Environment.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI();
            }

            app.UseResponseCompression();
            app.UseAuthorization();
            app.UseCors();


            app.MapControllers();

            app.Run();
        }
    }
}