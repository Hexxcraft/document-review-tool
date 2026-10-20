# Third-party dependencies

The original application code is separate from the third-party components below.
Package and upstream license terms govern those components; publication of this
portfolio does not relicense them.

| Component | Version referenced | Role | Upstream |
| --- | --- | --- | --- |
| Microsoft.Web.WebView2 | 1.0.4258.31 | Embedded browser and loader | [NuGet](https://www.nuget.org/packages/Microsoft.Web.WebView2/1.0.4258.31) |
| PDFtoImage | 5.4.0 | PDF page rasterization | [NuGet](https://www.nuget.org/packages/PDFtoImage/5.4.0) |
| UglyToad.PdfPig | 1.7.0-custom-5 | PDF text extraction | [NuGet](https://www.nuget.org/packages/UglyToad.PdfPig/1.7.0-custom-5) |

The build also restores transitive packages, including SkiaSharp and PDFium
components. Inspect package metadata and resolved dependencies before distributing
binaries:

```powershell
dotnet list .\HexxcorpDocumentReviewUtility.csproj package --include-transitive
```

The application embeds a WebView2 loader for its deployment fallback.
Microsoft Edge WebView2 Runtime and Microsoft Office are separate installed
products, not included in this source repository. Review their terms before
redistributing or deploying the application.

The banner and synthetic demonstration documents in this repository were created
for this portfolio. They do not contain real review records.
