using System.IO;
using System.Configuration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using ReportRenderer.Service;

namespace ReportRenderer.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Connection

            var username = ConfigurationManager.AppSettings["CRMUser"];
            var password = ConfigurationManager.AppSettings["CRMPassword"];
            var url = ConfigurationManager.AppSettings["CRMServiceURL"];
            var appId = ConfigurationManager.AppSettings["CRMAppId"];
            var redirectUri = ConfigurationManager.AppSettings["CRMRedirectURI"];

            var connectionString = "authtype=OAuth;" +
                $"Username={username};" +
                $"Password={password};" +
                $"Url={url};" +
                $"AppId={appId};" +
                $"RedirectUri={redirectUri};" +
                $"LoginPrompt=Never;";

            var serviceClient = new CrmServiceClient(connectionString);

            var service = serviceClient.OrganizationWebProxyClient
                ?? (IOrganizationService)serviceClient.OrganizationServiceProxy;

            if (service == null) return;

            #endregion

            var primaryEntity = new EntityReference("contact")
            {
                Id = new System.Guid("4E7AFFF9-A528-ED11-9DB1-000D3A1B4871")
            };

            var reportName = "Test";
            var format = RenderFormat.PDF;

            var bytes = ReportRendererService.RenderReport(reportName, format, primaryEntity, service);

            var filename = $@"{reportName.Replace(" ", "_")}{GetRenderFormatExtension(format)}";

            File.WriteAllBytes($@"..\..\..\ReportRenderer.Files\{filename}", bytes);
        }

        static string GetRenderFormatExtension(RenderFormat format)
        {
            switch (format)
            {
                case RenderFormat.DOCX:
                    return ".docx";
                case RenderFormat.PDF:
                    return ".pdf";
                case RenderFormat.TIFF:
                    return ".tiff";
                case RenderFormat.XLSX:
                    return ".xlsx";
                default:
                    return null;
            }
        }
    }
}
