# ReportRendererD365

## About the Project

This project consists of a main assembly [`ReportRenderer.Service`](./ReportRenderer.Service/ReportRendererService.cs), which contains a static fuction that allows for the rendering of SSRS reports in various contexts. The assembly is to be used with Microsoft's CRM Platform Dynamics 365, which includes SQL Server Reporting Services out of the box, but does not support rendering programatically through SDK nor API as of yet. The Dynamics 365 instance is passed to the static function as an `IOrganizationService`, which can be instantiated in various ways as seen in the other C# projects within the solution. The assembly also includes a `RenderFormat` enumeration to allow for the rendering of .docx, .pdf, .tiff, and .xls formats.

All other C# projects within the solution are various contexts that demonstrate the usage of `ReportRenderer.Service`, including a [Console Application](./ReportRenderer.Console/Program.cs) that can be hosted as an Azure WebJob, a Dynamics 365 [Plugin Assembly](./ReportRenderer.Plugins/EmailReport.cs) and [Custom Workflow Action Assembly](./ReportRenderer.Workflows/EmailReport.cs) complete with [unit tests](./ReportRenderer.Tests/), and an ASP .NET Web API with a single [`ReportController`](./ReportRenderer.Api/Controllers/ReportController.cs).

The remaining areas of the solution include `ReportRenderer.Files`, which contains an unmanaged Dynamics 365 solution containing the two assemblies and all dependencies, and `ReportRenderer.Reports`, which contains a single fetch-based SSRS report compatable with Dynamics 365 and used as the report definition to render in the unit tests.

## Getting Started

To build the solution, you will need

- Visual Studio 2019
- The [Microsoft Reporting Services Projects](https://marketplace.visualstudio.com/items?itemName=ProBITools.MicrosoftReportProjectsforVisualStudio) Extension
- The [Dynamics 365, version 9.0 Report Authoring Extension](https://www.microsoft.com/en-US/download/details.aspx?id=56973)
- SQL Server Data Tools (SSDT)
- .NET Framework 4.6.2 for the assemblies
- 4.7 for the colsole app and API

Once the repository is cloned and the solution is opened, rebuild the `ReportRenderer.Service` project and reinstall any NuGet dependencies if needed.

When the project builds without error, rebuild the entire solution and reinstall any NuGet Packages if needed.

Once the solution builds, the unit tests can be run on-demand to ensure everything is working. To begin using the Report Renderer with a live Dynamics 365 instance, simply add the service account credentials via App Registration or Username and Password to the Web.config and/or App.config for the desired projects. For the Plugin or Workflow, the `ReportRenderer_1_0_0_0.zip` file can be imported as an unmanaged solution to any v9 instance of Dynamics.

## Contact

For any questions, contact jacob.pendleton@outlook.com
