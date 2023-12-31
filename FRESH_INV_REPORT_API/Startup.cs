using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using FRESH_INV_REPORT_API.Services;

namespace FRESH_INV_REPORT_API
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {

            services.AddControllers();
            string connectionString = Configuration.GetConnectionString("OracleConnection");
            services.AddScoped<FreshReportServices>(provider=>
            {
                return new FreshReportServices(connectionString);
            });
            services.AddScoped<NetMovementServices>(provider =>
            {
                return new NetMovementServices(connectionString);
            });
            services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new OpenApiInfo { Title = "FRESH_INV_REPORT_API", Version = "v1" });
            });
            services.AddCors(options =>
            {
                options.AddPolicy("MyCorePolicy", builder =>
                {
                    builder.WithOrigins("http://192.168.51.26:86", "http://localhost:3000").AllowAnyMethod().AllowAnyHeader();
                });
            });
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
                app.UseSwagger();
                app.UseSwaggerUI(c => c.SwaggerEndpoint("/swagger/v1/swagger.json", "FRESH_INV_REPORT_API v1"));
            }

            app.UseHttpsRedirection();

            app.UseRouting();

            app.UseCors("MyCorePolicy");

            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
        }
    }
}
