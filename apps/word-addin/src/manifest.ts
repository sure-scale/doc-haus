import { WORD_ADDIN_ID, WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import { httpsOrigin } from "./client"
import { WORD_BOOTSTRAP_ORIGIN, wordBootstrapUrl } from "./bootstrap-contract"

const xml = (value: string) =>
  value.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")

/** The HTTPS bridge renders this with its configured, certificate-trusted origin. */
export function renderManifest(origin: string): string {
  const base = xml(`${WORD_BOOTSTRAP_ORIGIN}/word-addin`)
  const pane = xml(wordBootstrapUrl(origin))
  const domains = [...new Set([WORD_BOOTSTRAP_ORIGIN, httpsOrigin(origin)])]
  return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xsi:type="TaskPaneApp">
  <Id>${WORD_ADDIN_ID}</Id>
  <Version>${WORD_ADDIN_VERSION}</Version>
  <ProviderName>doc.haus</ProviderName>
  <DefaultLocale>en-GB</DefaultLocale>
  <DisplayName DefaultValue="doc.haus"/>
  <Description DefaultValue="Chat, review amendments and save your Word document to doc.haus."/>
  <IconUrl DefaultValue="${base}/assets/icon.png"/>
  <HighResolutionIconUrl DefaultValue="${base}/assets/icon.png"/>
  <SupportUrl DefaultValue="https://doc.haus"/>
  <AppDomains>${domains.map(domain => `<AppDomain>${xml(domain)}</AppDomain>`).join("")}</AppDomains>
  <Hosts><Host Name="Document"/></Hosts>
  <Requirements><Sets DefaultMinVersion="1.1"><Set Name="WordApi"/></Sets></Requirements>
  <DefaultSettings><SourceLocation DefaultValue="${pane}"/></DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements><bt:Sets DefaultMinVersion="1.1"><bt:Set Name="WordApi"/></bt:Sets></Requirements>
    <Hosts><Host xsi:type="Document"><Runtimes><Runtime resid="Pane.Url" lifetime="long"/></Runtimes><DesktopFormFactor>
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
      <bt:Urls><bt:Url id="Pane.Url" DefaultValue="${pane}"/></bt:Urls>
      <bt:ShortStrings><bt:String id="Group.Label" DefaultValue="doc.haus"/><bt:String id="Button.Label" DefaultValue="doc.haus"/></bt:ShortStrings>
      <bt:LongStrings><bt:String id="Button.Description" DefaultValue="Chat, review amendments, read matter documents and compare saved versions."/></bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`
}
