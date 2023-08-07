using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using ReportRenderer.Service;
using System;
using System.Data;
using System.Linq;

namespace ReportRenderer.Plugins
{
    public class EmailReport : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var service = ((IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory)))
                .CreateOrganizationService(context.UserId);

            var reportSend = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId,
                new ColumnSet("rpt_primaryentityname", "rpt_primaryentityid", "rpt_reportname", "rpt_format", "rpt_recipient"));

            var primaryEntityLogicalName = reportSend.GetAttributeValue<string>("rpt_primaryentityname");
            var primaryEntityId = new Guid(reportSend.GetAttributeValue<string>("rpt_primaryentityid"));

            var primaryEntity = service.Retrieve(primaryEntityLogicalName, primaryEntityId, new ColumnSet(false)).ToEntityReference();
            var primaryEntityName = primaryEntity.Name ?? "Report";

            var reportName = reportSend.GetAttributeValue<string>("rpt_reportname");
            var format = GetRenderFormatFromString(GetOptionSetLabel(reportSend, "rpt_format", service));
            var recipient = reportSend.GetAttributeValue<EntityReference>("rpt_recipient");

            var recipientEmail = service.Retrieve(recipient.LogicalName, recipient.Id, new ColumnSet("emailaddress1"))
                .GetAttributeValue<string>("emailaddress1");

            var senderEmail = service.Retrieve("systemuser", context.UserId, new ColumnSet("internalemailaddress"))
                .GetAttributeValue<string>("internalemailaddress");

            var filename = $@"{reportName.Replace(" ", "_")}{GetRenderFormatExtension(format)}";

            byte[] bytes;
            try
            {
                bytes = ReportRendererService.RenderReport(reportName, format, primaryEntity, service);
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.ToString());
            }

            #region Send Email

            var toparty = new Entity("activityparty")
            {
                ["addressused"] = recipientEmail,
                ["partyid"] = recipient,
            };

            var fromparty = new Entity("activityparty")
            {
                ["addressused"] = senderEmail,
                ["partyid"] = new EntityReference("systemuser", context.UserId)
            };

            var email = new Entity("email")
            {
                ["subject"] = $"{primaryEntityName} - {reportName}",
                ["description"] = $"Attached is your report.",
                ["to"] = new Entity[] { toparty },
                ["from"] = new Entity[] { fromparty }
            };
            email.Id = service.Create(email);

            var attachment = new Entity("activitymimeattachment")
            {
                ["objectid"] = email.ToEntityReference(),
                ["objecttypecode"] = email.LogicalName,
                ["filename"] = filename,
                ["mimetype"] = GetRenderFormatMimeType(format),
                ["body"] = Convert.ToBase64String(bytes)
            };
            service.Create(attachment);

            var sendRequest = new SendEmailRequest()
            {
                EmailId = email.Id,
                TrackingToken = "",
                IssueSend = true
            };
            var sendResponse = (SendEmailResponse)service.Execute(sendRequest);

            #endregion
        }

        string GetRenderFormatExtension(RenderFormat format)
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

        string GetRenderFormatMimeType(RenderFormat format)
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

        RenderFormat GetRenderFormatFromString(string format)
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

        string GetOptionSetLabel(Entity entity, string attribute, IOrganizationService service)
        {
            var attributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entity.LogicalName,
                LogicalName = attribute,
                RetrieveAsIfPublished = true
            };

            var attributeResponse = (RetrieveAttributeResponse)service.Execute(attributeRequest);
            var attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

            return attributeMetadata.OptionSet.Options
                .Where(x => x.Value == entity.GetAttributeValue<OptionSetValue>(attribute).Value)
                .Select(x => x.Label).FirstOrDefault().UserLocalizedLabel.Label;
        }
    }
}