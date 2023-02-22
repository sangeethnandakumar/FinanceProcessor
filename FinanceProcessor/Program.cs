using FinanceProcessor.Application;
using FinanceProcessor.Application.Handlers;
using FinanceProcessor.Infrastructure;
using FinanceProcessor.Infrastructure.Handlers;
using Microsoft.Extensions.DependencyInjection;
using Serilog;

namespace FinanceProcessor
{
    public class Startup
    {
        public void ConfigureServices(IServiceCollection services)
        {
            // Add your services here
            services.AddSingleton<IPDFHandler, PDFHandler>();
            services.AddSingleton<IExcelHandler, ExcelHandler>();
            services.AddSingleton<IStatementProcessor, StatementProcessor>();
            services.AddSingleton<Main>();
        }
    }

    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();           

            var services = new ServiceCollection();
            var startup = new Startup();

            // Add logging
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .Enrich.FromLogContext()
                .WriteTo.File(
                    "log.txt", 
                    rollingInterval: RollingInterval.Day, 
                    rollOnFileSizeLimit: true, 
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level}] {Message}{NewLine}{Exception}{Properties}{NewLine}"
                )
                .CreateLogger();

            services.AddLogging(builder =>
                builder.AddSerilog(Log.Logger, dispose: true));
            startup.ConfigureServices(services);

            var serviceProvider = services.BuildServiceProvider();
            var mainForm = serviceProvider.GetRequiredService<Main>();
            System.Windows.Forms.Application.Run(mainForm);
        }
    }
}