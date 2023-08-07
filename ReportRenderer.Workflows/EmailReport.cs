using System;
using System.Activities;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using ReportRenderer.Service;

namespace ReportRenderer.Workflows
{
    public class EmailReport : CodeActivity
    {
        [Input("Recipient")]
        [RequiredArgument]
        [ReferenceTarget("contact")]
        public InArgument<EntityReference> Recipient { get; set; }

        [Input("Report Name")]
        [RequiredArgument]
        public InArgument<string> ReportName { get; set; }

        [Input("Render Format")]
        [RequiredArgument]
        public InArgument<string> Format { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            var context = executionContext.GetExtension<IWorkflowContext>();
            var service = executionContext.GetExtension<IOrganizationServiceFactory>()
                .CreateOrganizationService(context.UserId);

            var primaryEntity = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(false)).ToEntityReference();

            var primaryEntityName = primaryEntity.Name ?? "Report";

            var recipient = Recipient.Get(executionContext);

            var recipientEmail = service.Retrieve(recipient.LogicalName, recipient.Id, new ColumnSet("emailaddress1"))
                .GetAttributeValue<string>("emailaddress1");

            var senderEmail = service.Retrieve("systemuser", context.UserId, new ColumnSet("internalemailaddress"))
                .GetAttributeValue<string>("internalemailaddress");

            var reportName = ReportName.Get(executionContext);
            var format = GetRenderFormatFromString(Format.Get(executionContext));
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
    }
}