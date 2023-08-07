using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using ReportRenderer.Service;
using System;
using System.Configuration;
using System.IO;
using System.Web.Http;
using ReportRenderer.Api.Utilities;

namespace ReportRenderer.Api.Controllers
{
    public partial class ReportController : ApiController
    {
        // https://localhost:44356/api/Report?LogicalName=contact&Id=4E7AFFF9-A528-ED11-9DB1-000D3A1B4871&ReportName=Test&Format=PDF
        [Route("api/Report")]
        public IHttpActionResult Get(string LogicalName, string Id, string ReportName, string Format)
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

            #endregion

            var primaryEntity = new EntityReference(LogicalName)
            {
                Id = new Guid(Id)
            };

            var format = Utils.GetRenderFormatFromString(Format);

            var bytes = ReportRendererService.RenderReport(ReportName, format, primaryEntity, service);

            return new HttpActionResultFile(new MemoryStream(bytes), Request, Utils.GetRenderFormatMimeType(format), $"{ReportName}{Utils.GetRenderFormatExtension(format)}");
        }
    }
}
