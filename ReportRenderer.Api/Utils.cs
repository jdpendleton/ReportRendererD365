using System.IO;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using ReportRenderer.Service;

namespace ReportRenderer.Api.Utilities
{
    public class HttpActionResultFile : IHttpActionResult
    {
        private readonly MemoryStream data;
        private readonly string filename;
        private readonly string format;
        private readonly HttpRequestMessage request;
        private HttpResponseMessage httpResponseMessage;

        public HttpActionResultFile(MemoryStream data, HttpRequestMessage request, string format, string filename)
        {
            this.data = data;
            this.request = request;
            this.filename = filename;
            this.format = format;
        }

        public System.Threading.Tasks.Task<HttpResponseMessage> ExecuteAsync(System.Threading.CancellationToken cancellationToken)
        {
            httpResponseMessage = request.CreateResponse(HttpStatusCode.OK);
            httpResponseMessage.Content = new StreamContent(data);
            httpResponseMessage.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
            httpResponseMessage.Content.Headers.ContentDisposition.FileName = filename;
            httpResponseMessage.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(format); // application/octet-stream
            return System.Threading.Tasks.Task.FromResult(httpResponseMessage);
        }
    }

    public class Utils
    {
        public static string GetRenderFormatExtension(RenderFormat format)
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

        public static string GetRenderFormatMimeType(RenderFormat format)
        {
            switch (format)
            {
                case RenderFormat.DOCX:
                    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                case RenderFormat.PDF:
                    return "application/pdf";
                case RenderFormat.TIFF:
                    return "image/tiff";
                case RenderFormat.XLSX:
                    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                default:
                    return null;
            }
        }

        public static RenderFormat GetRenderFormatFromString(string format)
        {
            switch (format)
            {
                case "DOCX":
                    return RenderFormat.DOCX;
                case "PDF":
                    return RenderFormat.PDF;
                case "TIFF":
                    return RenderFormat.TIFF;
                case "XLSX":
                    return RenderFormat.XLSX;
                default:
                    return RenderFormat.PDF;
            }
        }
    }
}
