const APP_ID = "52833864-41e4-4923-87a2-e2438c40a836";

function xmlText(value) {
  return String(value).replace(/[&<>"']/g, character => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;"
  })[character]);
}

export function buildManifest(baseAddress) {
  const parsed = new URL(baseAddress);
  if (parsed.protocol !== "https:" || parsed.username || parsed.password || parsed.search || parsed.hash) {
    throw new Error("Use the HTTPS directory address of the hosted test files.");
  }
  if (!parsed.pathname.endsWith("/")) parsed.pathname += "/";
  const base = parsed.href;
  const appPage = new URL("index.html", base);
  appPage.searchParams.set("v", "0.1.0");
  const page = xmlText(appPage.href);
  const image = size => xmlText(new URL(`icon-${size}.png`, base).href);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
 xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
 xsi:type="TaskPaneApp">
 <Id>${APP_ID}</Id>
 <Version>1.1.0.0</Version>
 <ProviderName>SDO Kalinga ICT</ProviderName>
 <DefaultLocale>en-US</DefaultLocale>
 <DisplayName DefaultValue="Item Analysis Generator"/>
 <Description DefaultValue="Generate subjects, grade paper answers, and calculate V14 item analysis."/>
 <IconUrl DefaultValue="${image(32)}"/>
 <HighResolutionIconUrl DefaultValue="${image(64)}"/>
 <SupportUrl DefaultValue="${page}"/>
 <AppDomains><AppDomain>${xmlText(parsed.origin)}</AppDomain></AppDomains>
 <Hosts><Host Name="Workbook"/></Hosts>
 <Requirements><Sets DefaultMinVersion="1.9"><Set Name="ExcelApi"/></Sets></Requirements>
 <DefaultSettings><SourceLocation DefaultValue="${page}"/></DefaultSettings>
 <Permissions>ReadWriteDocument</Permissions>
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
  <Hosts><Host xsi:type="Workbook"><DesktopFormFactor>
   <GetStarted><Title resid="GetStarted.Title"/><Description resid="GetStarted.Description"/><LearnMoreUrl resid="Taskpane.Url"/></GetStarted>
   <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <OfficeTab id="TabHome"><Group id="TIA.Hosting.Group">
     <Label resid="CommandsGroup.Label"/>
     <Icon><bt:Image size="16" resid="Icon.16"/><bt:Image size="32" resid="Icon.32"/><bt:Image size="80" resid="Icon.80"/></Icon>
     <Control xsi:type="Button" id="TIA.Hosting.Open">
      <Label resid="TaskpaneButton.Label"/>
      <Supertip><Title resid="TaskpaneButton.Label"/><Description resid="TaskpaneButton.Help"/></Supertip>
      <Icon><bt:Image size="16" resid="Icon.16"/><bt:Image size="32" resid="Icon.32"/><bt:Image size="80" resid="Icon.80"/></Icon>
      <Action xsi:type="ShowTaskpane"><TaskpaneId>TIA.Hosting.Pane</TaskpaneId><SourceLocation resid="Taskpane.Url"/></Action>
     </Control>
    </Group></OfficeTab>
   </ExtensionPoint>
  </DesktopFormFactor></Host></Hosts>
  <Resources>
   <bt:Images><bt:Image id="Icon.16" DefaultValue="${image(16)}"/><bt:Image id="Icon.32" DefaultValue="${image(32)}"/><bt:Image id="Icon.80" DefaultValue="${image(80)}"/></bt:Images>
   <bt:Urls><bt:Url id="Taskpane.Url" DefaultValue="${page}"/></bt:Urls>
   <bt:ShortStrings><bt:String id="GetStarted.Title" DefaultValue="Open Item Analysis Generator"/><bt:String id="CommandsGroup.Label" DefaultValue="Item Analysis"/><bt:String id="TaskpaneButton.Label" DefaultValue="Item Analysis Generator"/></bt:ShortStrings>
   <bt:LongStrings><bt:String id="GetStarted.Description" DefaultValue="On the Home tab, open Item Analysis Generator to prepare a workbook and generate a subject."/><bt:String id="TaskpaneButton.Help" DefaultValue="Generate a subject and check its paper answers."/></bt:LongStrings>
  </Resources>
 </VersionOverrides>
</OfficeApp>`;
}
