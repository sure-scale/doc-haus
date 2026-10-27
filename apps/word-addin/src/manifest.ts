import { WORD_ADDIN_ID, WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import { httpsOrigin } from "./client"

const xml = (value: string) =>
  value.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")

/** The HTTPS bridge renders this with its configured, certificate-trusted origin. */
export function renderManifest(origin: string): string {
  const base = xml(`${httpsOrigin(origin)}/word-addin`)
  return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xsi:type="TaskPaneApp">
  <Id>${WORD_ADDIN_ID}</Id>
  <Version>${WORD_ADDIN_VERSION}</Version>
  <ProviderName>doc.haus</ProviderName>
  <DefaultLocale>en-GB</DefaultLocale>
  <DisplayName DefaultValue="doc.haus"/>
  <Description DefaultValue="Save your Word document back to its matter in doc.haus."/>
  <IconUrl DefaultValue="${base}/assets/icon.png"/>
  <HighResolutionIconUrl DefaultValue="${base}/assets/icon.png"/>
  <SupportUrl DefaultValue="https://doc.haus"/>
  <AppDomains><AppDomain>${xml(httpsOrigin(origin))}</AppDomain></AppDomains>
  <Hosts><Host Name="Document"/></Hosts>
  <Requirements><Sets DefaultMinVersion="1.1"><Set Name="WordApi"/></Sets></Requirements>
  <DefaultSettings><SourceLocation DefaultValue="${base}/pane.html?version=${WORD_ADDIN_VERSION}"/></DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements><bt:Sets DefaultMinVersion="1.1"><bt:Set Name="WordApi"/></bt:Sets></Requirements>
    <Hosts><Host xsi:type="Document"><DesktopFormFactor>
      <ExtensionPoint xsi:type="PrimaryCommandSurface"><OfficeTab id="TabHome"><Group id="Dochaus.Group">
        <Label resid="Group.Label"/>
        <Icon><bt:Image size="16" resid="Icon"/><bt:Image size="32" resid="Icon"/><bt:Image size="80" resid="Icon"/></Icon>
        <Control xsi:type="Button" id="Dochaus.Open">
          <Label resid="Button.Label"/>
          <Supertip><Title resid="Button.Label"/><Description resid="Button.Description"/></Supertip>
          <Icon><bt:Image size="16" resid="Icon"/><bt:Image size="32" resid="Icon"/><bt:Image size="80" resid="Icon"/></Icon>
          <Action xsi:type="ShowTaskpane"><TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId><SourceLocation resid="Pane.Url"/></Action>
        </Control>
      </Group></OfficeTab></ExtensionPoint>
    </DesktopFormFactor></Host></Hosts>
    <Resources>
      <bt:Images><bt:Image id="Icon" DefaultValue="${base}/assets/icon.png"/></bt:Images>
      <bt:Urls><bt:Url id="Pane.Url" DefaultValue="${base}/pane.html?version=${WORD_ADDIN_VERSION}"/></bt:Urls>
      <bt:ShortStrings><bt:String id="Group.Label" DefaultValue="doc.haus"/><bt:String id="Button.Label" DefaultValue="Save to doc.haus"/></bt:ShortStrings>
      <bt:LongStrings><bt:String id="Button.Description" DefaultValue="Open the doc.haus pane to save this document back to its matter."/></bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`
}
