# Modernization Summary Report

Date  |  Duration  |  Source Page  |  Target Page Url  |  Status
-------------  |  -------------  |  -------------  |  -------------  |  -------------
11-02-2026 20:39:03  |  00:00:04  |  [/Pages/Covid-Tracker.aspx](https://caje77sharepoint.sharepoint.com/sites/CajeTech/Pages/Covid-Tracker.aspx)  |  [/SitePages/Covid-Tracker.aspx](https://caje77sharepoint.sharepoint.com/sites/CajIntra/SitePages/Covid-Tracker.aspx)  |  Successful with 0 warnings and 1 non-critical errors
## Errors during transformation

Date  |  Source Page  |  Operation  |  Message
-------------  |  -------------  |  -------------  |  -------------
11-02-2026 20:39:03  |  /Pages/Covid-Tracker.aspx  |   Page Creation  |  Checking Page Exists    at System.Guid.GuidResult.SetFailure(ParseFailure failureKind)
   at System.Guid.TryParseGuid(ReadOnlySpan`1 guidString, GuidResult& result)
   at System.Guid.Parse(String input)
   at PnP.Framework.Modernization.Publishing.PublishingPageTransformator.Load(ClientContext sourceContext, ClientContext targetContext, PublishingPageTransformationInformation publishingPageTransformationInformation, List& pagesLibrary)
   at PnP.Framework.Modernization.Publishing.PublishingPageTransformator.Transform(PublishingPageTransformationInformation publishingPageTransformationInformation)

