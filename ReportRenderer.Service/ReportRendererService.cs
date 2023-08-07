using Microsoft.Crm.Sdk.Messages;
using Microsoft.Reporting.WebForms;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml;

namespace ReportRenderer.Service
{
    public class ReportRendererService
    {
        public static byte[] RenderReport(string reportName, RenderFormat format, EntityReference primaryEntity, IOrganizationService service)
        {
            #region Get Report Definitions

            var reportQuery = new QueryByAttribute("report")
            {
                ColumnSet = new ColumnSet("name", "reportid")
            };

            reportQuery.AddAttributeValue("name", reportName);

            var report = service.RetrieveMultiple(reportQuery).Entities.FirstOrDefault();

            var rdlRequest = new DownloadReportDefinitionRequest
            {
                ReportId = report.GetAttributeValue<Guid>("reportid")
            };

            var rdlResponse = (DownloadReportDefinitionResponse)service.Execute(rdlRequest);

            var reportDefinitions = new List<ReportDefinition>();

            var document = new XmlDocument();
            document.LoadXml(rdlResponse.BodyText);

            var ns = new XmlNamespaceManager(document.NameTable);

            ns.AddNamespace("rd", "http://schemas.microsoft.com/SQLServer/reporting/reportdesigner");
            ns.AddNamespace("r", "http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition");

            var ds = document.SelectNodes(@"/*/r:DataSets", ns);

            if (ds.Count == 0)
            {
                ds = document.GetElementsByTagName("DataSets");
            }

            #endregion

            #region Add Filtering

            foreach (XmlNode node in ds[0].ChildNodes)
            {
                var commandText = node.ChildNodes.OfType<XmlElement>().Where(e => e.Name == "Query").FirstOrDefault()
                    .ChildNodes.OfType<XmlElement>().Where(c => c.Name == "CommandText").FirstOrDefault();

                var fetch = new XmlDocument();
                fetch.LoadXml(commandText.InnerText);

                var entity = fetch.SelectSingleNode("/fetch").ChildNodes.OfType<XmlElement>()
                    .Where(e => e.GetAttribute("Name") == primaryEntity.LogicalName).FirstOrDefault();

                if (entity != null)
                {
                    var filterNode = entity.ChildNodes[0].ChildNodes.OfType<XmlElement>()
                        .Where(e => e.Name == "filter").FirstOrDefault();

                    if (filterNode == null)
                    {
                        var filterElem = fetch.CreateElement("filter");
                        filterElem.SetAttribute("type", "and");

                        var conditionElem = fetch.CreateElement("condition");
                        conditionElem.SetAttribute("attribute", $"{primaryEntity.LogicalName}id");
                        conditionElem.SetAttribute("operator", "eq");
                        conditionElem.SetAttribute("value", primaryEntity.Id.ToString());

                        filterElem.AppendChild(conditionElem);
                        entity.AppendChild(filterElem);
                    }
                    else
                    {
                        var existingFilter = fetch.SelectSingleNode("/fetch//entity//filter");

                        var conditionElem = fetch.CreateElement("condition");
                        conditionElem.SetAttribute("attribute", $"{primaryEntity.LogicalName}id");
                        conditionElem.SetAttribute("operator", "eq");
                        conditionElem.SetAttribute("value", primaryEntity.Id.ToString());

                        existingFilter.AppendChild(conditionElem);
                    }
                }

                reportDefinitions.Add(new ReportDefinition()
                {
                    DataSetName = node.Attributes["Name"].Value,
                    FetchXML = fetch.InnerXml.ToString()
                });
            }

            #endregion

            #region Add Data Sources

            var viewer = new ReportViewer()
            {
                ProcessingMode = ProcessingMode.Local
            };

            using (var textReader = new StringReader(rdlResponse.BodyText))
            {
                viewer.LocalReport.LoadReportDefinition(textReader);
            };

            viewer.LocalReport.EnableHyperlinks = true;

            foreach (var reportDefinition in reportDefinitions)
            {
                var response = service.RetrieveMultiple(new FetchExpression(reportDefinition.FetchXML));

                var resultsTable = new DataTable("Dataset1");
                if (response.Entities.Count > 0)
                {
                    for (int i = 0; i < response.Entities.Count; i++)
                    {
                        var entity = response.Entities[i];
                        var row = resultsTable.NewRow();
                        foreach (var attribute in entity.Attributes)
                        {
                            if (!resultsTable.Columns.Contains(attribute.Key))
                            {
                                resultsTable.Columns.Add(attribute.Key);
                            }

                            if (GetAttributeValue(attribute.Value) == null)
                            {
                                row[attribute.Key] = string.Empty;
                            }
                            else
                            {
                                row[attribute.Key] = GetAttributeValue(attribute.Value).ToString();
                            }

                        }
                        foreach (var fv in entity.FormattedValues)
                        {
                            if (!resultsTable.Columns.Contains(fv.Key + "name"))
                            {
                                resultsTable.Columns.Add(fv.Key + "name");
                            }
                            row[fv.Key + "name"] = fv.Value;
                        }

                        resultsTable.Rows.Add(row);
                    }
                }

                var rds = new ReportDataSource(reportDefinition.DataSetName, resultsTable);
                viewer.LocalReport.DataSources.Add(rds);
            }

            #endregion

            return viewer.LocalReport.Render(GetRenderFormatString(format));
        }

        private static string GetRenderFormatString(object format)
        {
            switch (format)
            {
                case RenderFormat.DOCX:
                    return "WORDOPENXML";
                case RenderFormat.PDF:
                    return "PDF";
                case RenderFormat.TIFF:
                    return "IMAGE";
                case RenderFormat.XLSX:
                    return "EXCELOPENXML";
                default:
                    return null;
            }
        }

        private static object GetAttributeValue(object entityValue)
        {
            switch (entityValue.ToString())
            {
                case "Microsoft.Xrm.Sdk.EntityReference":
                    return ((EntityReference)entityValue).Name;
                case "Microsoft.Xrm.Sdk.OptionSetValue":
                    return ((OptionSetValue)entityValue).Value.ToString();
                case "Microsoft.Xrm.Sdk.Money":
                    return ((Money)entityValue).Value.ToString();
                case "Microsoft.Xrm.Sdk.AliasedValue":
                    return GetAttributeValue(((AliasedValue)entityValue).Value);
                default:
                    return entityValue.ToString();
            }
        }
    }

    internal class ReportDefinition
    {
        public string DataSetName { get; internal set; }
        public string FetchXML { get; internal set; }
    }

    public enum RenderFormat
    {
        DOCX,
        PDF,
        TIFF,
        XLSX
    }
}
