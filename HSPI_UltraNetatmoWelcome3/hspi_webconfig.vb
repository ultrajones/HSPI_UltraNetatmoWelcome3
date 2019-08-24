Imports System.Text
Imports System.Web
Imports Scheduler
Imports HomeSeerAPI
Imports System.Collections.Specialized
Imports System.Web.UI.WebControls
Imports System.IO

Public Class hspi_webconfig
  Inherits clsPageBuilder

  Public hspiref As HSPI

  Dim TimerEnabled As Boolean

  ''' <summary>
  ''' Initializes new webconfig
  ''' </summary>
  ''' <param name="pagename"></param>
  ''' <remarks></remarks>
  Public Sub New(ByVal pagename As String)
    MyBase.New(pagename)
  End Sub

#Region "Page Building"

  ''' <summary>
  ''' Web pages that use the clsPageBuilder class and registered with hs.RegisterLink and hs.RegisterConfigLink will then be called through this function. 
  ''' A complete page needs to be created and returned.
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <param name="queryString"></param>
  ''' <param name="instance"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String, instance As String) As String

    Try

      Dim stb As New StringBuilder

      '
      ' Called from the start of your page to reset all internal data structures in the clsPageBuilder class, such as menus.
      '
      Me.reset()

      '
      ' Determine if user is authorized to access the web page
      '
      Dim LoggedInUser As String = hs.WEBLoggedInUser()
      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      '
      ' Handle any queries like mode=something
      '
      Dim parts As Collections.Specialized.NameValueCollection = Nothing
      If (queryString <> "") Then
        parts = HttpUtility.ParseQueryString(queryString)
      End If

      Dim Header As New StringBuilder
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultranetatmowelcome3/js/lightbox.min.js""></script>")

      Header.AppendLine("<link type=""text/css"" rel=""stylesheet"" href=""/hspi_ultranetatmowelcome3/css/lightbox.min.css"" />")
      Me.AddHeader(Header.ToString)

      Dim pageTile As String = String.Format("{0} {1}", pageName, instance).TrimEnd
      stb.Append(hs.GetPageHeader(pageName, pageTile, "", "", False, False))

      '
      ' Start the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

      '
      ' A message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
      '
      stb.Append(clsPageBuilder.DivStart("divErrorMessage", "class='errormessage'"))
      stb.Append(clsPageBuilder.DivEnd)

      Me.RefreshIntervalMilliSeconds = 3000
      stb.Append(Me.AddAjaxHandlerPost("id=timer", pageName))

      If WEBUserIsAuthorized(LoggedInUser, USER_ROLES_AUTHORIZED) = False Then
        '
        ' Current user not authorized
        '
        stb.Append(WebUserNotUnauthorized(LoggedInUser))
      Else
        '
        ' Specific page starts here
        '
        stb.Append(BuildContent)
      End If

      '
      ' End the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivEnd)

      '
      ' Add the body html to the page
      '
      Me.AddBody(stb.ToString)

      '
      ' Return the full page
      '
      Return Me.BuildPage()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "GetPagePlugin")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the HTML content
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildContent() As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table border='0' cellpadding='0' cellspacing='0' width='1000'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td width='1000' align='center' style='color:#FF0000; font-size:14pt; height:30px;'><strong><div id='divMessage'>&nbsp;</div></strong></td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", BuildTabs())
      stb.AppendLine(" </tr>")
      stb.AppendLine("</table>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildContent")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the jQuery Tabss
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function BuildTabs() As String

    Try

      Dim stb As New StringBuilder
      Dim tabs As clsJQuery.jqTabs = New clsJQuery.jqTabs("oTabs", Me.PageName)
      Dim tab As New clsJQuery.Tab

      tabs.postOnTabClick = True

      tab.tabTitle = "Status"
      tab.tabDIVID = "tabStatus"
      tab.tabContent = "<div id='divStatus'>" & BuildTabStatus() & "</div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Options"
      tab.tabDIVID = "tabOptions"
      tab.tabContent = "<div id='divOptions'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Cameras"
      tab.tabDIVID = "tabCameras"
      tab.tabContent = "<div id='divCameras'>" & BuildTabCameras() & "</div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Events"
      tab.tabDIVID = "tabEvents"
      tab.tabContent = "<div id='divEvents'>" & BuildTabEvents() & "</div>"
      tabs.tabs.Add(tab)

      Return tabs.Build

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabs")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Status Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabStatus(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine(clsPageBuilder.FormStart("frmStatus", "frmStatus", "Post"))

      stb.AppendLine("<div>")
      stb.AppendLine("<table>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Plug-In Status </legend>")
      stb.AppendLine("     <table style=""width: 100%"">")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Name:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", IFACE_NAME)
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Status:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", "OK")
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Version:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", HSPI.Version)
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Netatmo Authentication Status </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Authenticated:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", NetatmoAPI.ValidAPIToken())
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Success:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", NetatmoAPI.QuerySuccessCount())
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Failure:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", NetatmoAPI.QueryFailureCount())
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table>")
      stb.AppendLine("</div>")

      stb.AppendLine(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divStatus", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabStatus")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Options Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabOptions(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table cellspacing='0' width='100%'>")

      stb.Append(clsPageBuilder.FormStart("frmOptions", "frmOptions", "Post"))

      '
      ' General Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Netatmo Credentials</td>")
      stb.AppendLine(" </tr>")

      Dim tbUsername As New clsJQuery.jqTextBox("NetatmoUsername", "text", gAPIUsername, PageName, 30, False)
      tbUsername.id = "NetatmoUsername"
      tbUsername.promptText = "Enter your Netatmo Username"
      tbUsername.toolTip = tbUsername.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Username</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}&nbsp;Required</td>{1}", tbUsername.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      Dim tbPassword As New clsJQuery.jqTextBox("NetatmoPassword", "text", "", PageName, 30, False)
      tbPassword.id = "NetatmoPassword"
      tbPassword.promptText = "Enter your Netatmo Password (the password field will be emtpy after a page refresh)"
      tbPassword.toolTip = tbPassword.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Password</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}&nbsp;Required</td>{1}", tbPassword.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Netatmo Client</td>")
      stb.AppendLine(" </tr>")

      Dim txtAPIClientId As String = GetSetting("API", "ClientId", gAPIClientId)
      Dim tbAPIClientId As New clsJQuery.jqTextBox("APIClientId", "text", txtAPIClientId, PageName, 50, False)
      tbAPIClientId.id = "APIClientId"
      tbAPIClientId.promptText = "Enter your Netatmo Application Id (visit https://dev.netatmo.com/dev/createapp to obtain your API key)."
      tbAPIClientId.toolTip = tbAPIClientId.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Client&nbsp;Id</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}&nbsp;Required</td>{1}", tbAPIClientId.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      Dim txtAPIClientSecret As String = GetSetting("API", "APIClientSecret", "")
      Dim tbAPIClientSecret As New clsJQuery.jqTextBox("APIClientSecret", "text", txtAPIClientSecret, PageName, 50, False)
      tbAPIClientSecret.id = "APIClientSecret"
      tbAPIClientSecret.promptText = "Enter your Netatmo Application secret (the secret field will be emtpy after a page refresh)."
      tbAPIClientSecret.toolTip = tbAPIClientSecret.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Client&nbsp;Secret</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}&nbsp;Required</td>{1}", tbAPIClientSecret.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Netatmo Options</td>")
      stb.AppendLine(" </tr>")

      Dim selUnitType As New clsJQuery.jqDropList("selUnitType", Me.PageName, False)
      selUnitType.id = "selUnitType"
      selUnitType.toolTip = "The format used to display temperatures, rainfall and barometric pressure.  The default format is U.S customary units."

      Dim strUnitType As String = GetSetting("Options", "UnitType", "0")
      selUnitType.AddItem("U.S. customary units (miles, °F, etc...)", "0", strUnitType = "0")
      selUnitType.AddItem("Metric system units (kms, °C, etc...)", "1", strUnitType = "1")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Unit&nbsp;Type</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selUnitType.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      Dim selUpdateFrequency As New clsJQuery.jqDropList("selUpdateFrequency", Me.PageName, False)
      selUpdateFrequency.id = "selUpdateFrequency"
      selUpdateFrequency.toolTip = "Specify how often the plug-in should read values from your Netatmo Welcome."

      Dim txtUpdateFrequency As String = GetSetting("Options", "WelcomeUpdate", gWelcomeUpdate)
      For index As Integer = 1 To 60 Step 1
        Dim units As String = IIf(index = 1, "Minute", "Minutes")
        Dim value As String = index.ToString
        Dim desc As String = String.Format("{0} {1}", index.ToString, units)
        selUpdateFrequency.AddItem(desc, value, index.ToString = txtUpdateFrequency)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Welcome&nbsp;Update</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selUpdateFrequency.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Netatmo Options (Downloaded Video Quality)
      '
      Dim selVideoQuality As New clsJQuery.jqDropList("selVideoQuality", Me.PageName, False)
      selVideoQuality.id = "selVideoQuality"
      selVideoQuality.toolTip = "Specify the Video Quality of the downloaded Netatmo Welcome events."

      Dim strVideoQuality As String = GetSetting("Options", "VideoQuality", gEventVideoQuality)
      selVideoQuality.AddItem("Disabled", "disabled", strVideoQuality = "disabled")
      'selVideoQuality.AddItem("Poor 640x360", "poor", strVideoQuality = "poor")
      selVideoQuality.AddItem("Low 640x360", "low", strVideoQuality = "low")
      selVideoQuality.AddItem("Medium 1280x720", "medium", strVideoQuality = "medium")
      selVideoQuality.AddItem("High 1920x1080", "high", strVideoQuality = "high")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Downloaded&nbsp;Video&nbsp;Quality</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selVideoQuality.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Web Page Access
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Web Page Access</td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Authorized User Roles</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", BuildWebPageAccessCheckBoxes, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Application Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Application Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' Application Logging Level
      '
      Dim selLogLevel As New clsJQuery.jqDropList("selLogLevel", Me.PageName, False)
      selLogLevel.id = "selLogLevel"
      selLogLevel.toolTip = "Specifies the plug-in logging level."

      Dim itemValues As Array = System.Enum.GetValues(GetType(LogLevel))
      Dim itemNames As Array = System.Enum.GetNames(GetType(LogLevel))

      For i As Integer = 0 To itemNames.Length - 1
        Dim itemSelected As Boolean = IIf(gLogLevel = itemValues(i), True, False)
        selLogLevel.AddItem(itemNames(i), itemValues(i), itemSelected)
      Next
      selLogLevel.autoPostBack = True

      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", "Logging&nbsp;Level", vbCrLf)
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selLogLevel.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divOptions", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabOptions")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Options Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabCameras(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder
      Dim cameras As New ArrayList

      stb.Append(clsPageBuilder.FormStart("frmCameras", "frmCameras", "Post"))

      Dim selSnapshotsMaxWidth As New clsJQuery.jqDropList("selSnapshotsMaxWidth", Me.PageName, False)

      selSnapshotsMaxWidth.id = "selSnapshotsMaxWidth"
      selSnapshotsMaxWidth.toolTip = "Specifies the maximum snapshot width."
      selSnapshotsMaxWidth.AddItem("Auto", "Auto", IIf(gSnapshotMaxWidth = "Auto", True, False))
      selSnapshotsMaxWidth.AddItem("160 px", "160px", IIf(gSnapshotMaxWidth = "160px", True, False))
      selSnapshotsMaxWidth.AddItem("320 px", "320px", IIf(gSnapshotMaxWidth = "320px", True, False))
      selSnapshotsMaxWidth.AddItem("640 px", "640px", IIf(gSnapshotMaxWidth = "640px", True, False))
      selSnapshotsMaxWidth.autoPostBack = True

      '
      ' General Options
      '
      stb.AppendLine("<table cellspacing='0' width='100%'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Welcome Cameras</td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'><div id='lastRefresh'/></td>")
      stb.AppendLine("  <td class='tablecell' align='right'>Snapshot Width: " & selSnapshotsMaxWidth.Build & "</td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' colspan='2'>")

      Dim WelcomeCameras As List(Of hspi_netatmo_api.Camera) = NetatmoAPI.GetWelcomeCameras
      For Each WelcomeCamera As hspi_netatmo_api.Camera In WelcomeCameras
        Dim camera_id As String = WelcomeCamera.id
        Dim dv_addr As String = String.Format("camera-{0}", camera_id.Replace(":", "-"))
        Dim strSnapshotFilename As String = String.Format("{0}.jpg", dv_addr.ToLower)

        cameras.Add(dv_addr)

        stb.AppendFormat("  <div class=""tablecell"" style=""float:left; border: solid black 1px; margin:2px; text-align:center; width:{0};"">", gSnapshotMaxWidth)
        stb.AppendFormat("   <a id=""lnk_{0}"" href=""#"" title=""{1}"" data-lightbox=""lightbox[0]"">", dv_addr, WelcomeCamera.name)
        stb.AppendFormat("    <img id=""img_{0}"" rel=""lightbox[0]"" style=""width:100%"" />", dv_addr)
        stb.AppendLine("   </a>")
        stb.AppendFormat("   <div>{0}</div><div>{1}</div>", WelcomeCamera.name, dv_addr)
        stb.AppendLine("  </div>")

      Next

      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine("</table>")

      '
      ' Update the Refresh Interval
      '
      Dim iRefreshInterval As Integer = 1000
      Dim strRefreshInterval As String = GetSetting("Options", "WelcomeUpdate", "5")
      If IsNumeric(strRefreshInterval) = True Then
        iRefreshInterval *= Integer.Parse(strRefreshInterval)
      End If
      iRefreshInterval += 10000

      stb.AppendLine("<script>")
      stb.AppendLine("function refreshSnapshots() {")
      stb.AppendLine("  var ticks = new Date().getTime();")
      stb.AppendLine("  var url = 'images/hspi_ultranetatmowelcome3/cameras/';")
      For Each dv_addr As String In cameras
        Dim strSnapshotFilename As String = String.Format("{0}.jpg", dv_addr)
        stb.AppendLine("    $('#img_" & dv_addr & "').attr('src', url + '" & strSnapshotFilename & "?ticks=' + ticks);")
        stb.AppendLine("    $('#lnk_" & dv_addr & "').attr('href', url + '" & strSnapshotFilename & "?ticks=' + ticks);")
      Next
      stb.AppendLine("    $('#lastRefresh').html( new Date() + '' );")
      stb.AppendLine("};")

      stb.AppendLine("$(function() { refreshSnapshots(); setInterval(function() { refreshSnapshots(); }, " & iRefreshInterval.ToString & ");});")

      'stb.AppendLine("var options = {};")
      'stb.AppendLine("$( "".effect"" ).effect( ""highlight"", options, 3000, callback );")
      'stb.AppendLine("});")
      'stb.AppendLine("function callback() { };")
      stb.AppendLine("</script>")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divCameras", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabCameras")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Options Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabEvents(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder
      Dim cameras As New ArrayList

      stb.Append(clsPageBuilder.FormStart("frmEvents", "frmEvents", "Post"))

      Dim selSnapshotEventMaxWidth As New clsJQuery.jqDropList("selSnapshotEventMaxWidth", Me.PageName, False)

      selSnapshotEventMaxWidth.id = "selSnapshotEventMaxWidth"
      selSnapshotEventMaxWidth.toolTip = "Specifies the maximum snapshot event width."
      'selSnapshotEventMaxWidth.AddItem("Auto", "Auto", IIf(gSnapshotMaxWidth = "Auto", True, False))
      selSnapshotEventMaxWidth.AddItem("160 px", "160px", IIf(gSnapshotEventMaxWidth = "160px", True, False))
      selSnapshotEventMaxWidth.AddItem("320 px", "320px", IIf(gSnapshotEventMaxWidth = "320px", True, False))
      selSnapshotEventMaxWidth.AddItem("640 px", "640px", IIf(gSnapshotEventMaxWidth = "640px", True, False))
      selSnapshotEventMaxWidth.autoPostBack = True

      '
      ' General Options
      '
      stb.AppendLine("<table cellspacing='0' width='100%'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Welcome Events</td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>&nbsp;</td>")
      stb.AppendLine("  <td class='tablecell' align='right'>Snapshot Width: " & selSnapshotEventMaxWidth.Build & "</td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' colspan='2'>")

      Dim WelcomeCameraNames As New Specialized.StringDictionary

      Dim WelcomeCameras As List(Of hspi_netatmo_api.Camera) = NetatmoAPI.GetWelcomeCameras
      For Each WelcomeCamera As hspi_netatmo_api.Camera In WelcomeCameras
        If WelcomeCameraNames.ContainsKey(WelcomeCamera.id) = False Then
          WelcomeCameraNames.Add(WelcomeCamera.id, WelcomeCamera.name)
        End If
      Next

      Dim WelcomeEvents As List(Of hspi_netatmo_api.Event) = NetatmoAPI.GetWelcomeEvents
      For Each [Event] As hspi_netatmo_api.Event In WelcomeEvents.OrderBy(Function(x) x.time)
        Dim event_id As String = [Event].id
        Dim netcam_id As String = [Event].camera_id
        Dim netcam_name As String = "Unknown Camera"
        Dim event_date As String = ConvertEpochToDateTime([Event].time).ToString
        Dim event_name As String = [Event].message

        If WelcomeCameraNames.ContainsKey(netcam_id) = True Then
          netcam_name = WelcomeCameraNames(netcam_id)
        End If

        Dim video_download As Boolean = [Event].video_downloaded

        If video_download = True Then

          Dim EventDirectory As String = "/images/hspi_ultranetatmowelcome3/events"

          If [Event].type = "outdoor" Then

            For Each Outdoor_Event As hspi_netatmo_api.Outdoor_Event In [Event].event_list

              Dim snapshot_filename As String = String.Format("{0}/{1}/{2}.jpg", EventDirectory, [Event].id.ToLower, Outdoor_Event.id.ToLower)
              Dim video_filename As String = String.Format("{0}/{1}/index.html", EventDirectory, Outdoor_Event.id.ToLower)

              If Outdoor_Event.snapshot.filename.Length = 0 Then
                snapshot_filename = String.Format("{0}/{1}/{2}.jpg", EventDirectory, [Event].id.ToLower, Outdoor_Event.snapshot.id.ToLower)
              End If

              Dim outdoor_event_name As String = Outdoor_Event.message
              Dim outdoor_event_date As String = ConvertEpochToDateTime(Outdoor_Event.time).ToString

              If gNetatmoCameras.Any(Function(s) s.id = [Event].camera_id) = True Then
                Dim NetatmoCamera As hspi_netatmo_api.Camera = gNetatmoCameras.Find(Function(s) s.id = [Event].camera_id)
                Dim camera_url As String = NetatmoCamera.camera_url
                Dim camera_vpn_url As String = NetatmoCamera.vpn_url
                Dim camera_is_local As String = NetatmoCamera.is_local

                video_filename = String.Format("{0}/vod/{1}/index.m3u8", camera_vpn_url, [Event].id.ToLower)
              End If

              stb.AppendFormat("  <div class=""tablecell"" style=""float:left; border: solid black 1px; margin:2px; text-align:center; width:{0};"">", gSnapshotEventMaxWidth)
              stb.AppendFormat("   <a id=""lnk_{0}"" href=""{1}"" title=""{2}"" data-lightbox=""lightbox[1]"">", event_id, snapshot_filename, netcam_name)
              stb.AppendFormat("    <img id=""img_{0}"" rel=""data-lightbox[1]"" style=""width:100%"" src='{1}' />", event_id, snapshot_filename)
              stb.AppendFormat("   </a>")
              stb.AppendFormat("   <div>{0}</div><div>{1}</div><div>{2}</div>", netcam_name, outdoor_event_date, outdoor_event_name)

              stb.AppendLine("  </div>")

            Next

          Else

            Dim snapshot_filename As String = String.Format("{0}/{1}/{2}.jpg", EventDirectory, [Event].id.ToLower, [Event].snapshot.id.ToLower)
            Dim video_filename As String = String.Format("{0}/{1}/index.html", EventDirectory, [Event].id.ToLower)

            If gNetatmoCameras.Any(Function(s) s.id = [Event].camera_id) = True Then
              Dim NetatmoCamera As hspi_netatmo_api.Camera = gNetatmoCameras.Find(Function(s) s.id = [Event].camera_id)
              Dim camera_url As String = NetatmoCamera.camera_url
              Dim camera_vpn_url As String = NetatmoCamera.vpn_url
              Dim camera_is_local As String = NetatmoCamera.is_local

              video_filename = String.Format("{0}/vod/{1}/index.m3u8", camera_vpn_url, [Event].id.ToLower)
            End If

            stb.AppendFormat("  <div class=""tablecell"" style=""float:left; border: solid black 1px; margin:2px; text-align:center; width:{0};"">", gSnapshotEventMaxWidth)
            stb.AppendFormat("   <a id=""lnk_{0}"" href=""{1}"" title=""{2}"" data-lightbox=""lightbox[1]"">", event_id, snapshot_filename, netcam_name)
            stb.AppendFormat("    <img id=""img_{0}"" rel=""data-lightbox[1]"" style=""width:100%"" src='{1}' />", event_id, snapshot_filename)
            stb.AppendFormat("   </a>")
            stb.AppendFormat("   <div>{0}</div><div>{1}</div><div>{2}</div>", netcam_name, event_date, event_name)

            stb.AppendLine("  </div>")

          End If

        End If

      Next

      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine("</table>")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divEvents", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabEvents")
      Return "Error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Web Page Access Checkbox List
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function BuildWebPageAccessCheckBoxes()

    Try

      Dim stb As New StringBuilder

      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      Dim cb1 As New clsJQuery.jqCheckBox("chkWebPageAccess_Guest", "Guest", Me.PageName, True, True)
      Dim cb2 As New clsJQuery.jqCheckBox("chkWebPageAccess_Admin", "Admin", Me.PageName, True, True)
      Dim cb3 As New clsJQuery.jqCheckBox("chkWebPageAccess_Normal", "Normal", Me.PageName, True, True)
      Dim cb4 As New clsJQuery.jqCheckBox("chkWebPageAccess_Local", "Local", Me.PageName, True, True)

      cb1.id = "WebPageAccess_Guest"
      cb1.checked = CBool(USER_ROLES_AUTHORIZED And USER_GUEST)

      cb2.id = "WebPageAccess_Admin"
      cb2.checked = CBool(USER_ROLES_AUTHORIZED And USER_ADMIN)
      cb2.enabled = False

      cb3.id = "WebPageAccess_Normal"
      cb3.checked = CBool(USER_ROLES_AUTHORIZED And USER_NORMAL)

      cb4.id = "WebPageAccess_Local"
      cb4.checked = CBool(USER_ROLES_AUTHORIZED And USER_LOCAL)

      stb.Append(clsPageBuilder.FormStart("frmWebPageAccess", "frmWebPageAccess", "Post"))

      stb.Append(cb1.Build())
      stb.Append(cb2.Build())
      stb.Append(cb3.Build())
      stb.Append(cb4.Build())

      stb.Append(clsPageBuilder.FormEnd())

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildWebPageAccessCheckBoxes")
      Return "Error - " & Err.Description
    End Try

  End Function

#End Region

#Region "Page Processing"

  ''' <summary>
  ''' Post a message to this web page
  ''' </summary>
  ''' <param name="sMessage"></param>
  ''' <remarks></remarks>
  Sub PostMessage(ByVal sMessage As String)

    Try

      Me.divToUpdate.Add("divMessage", sMessage)

      Me.pageCommands.Add("starttimer", "")

      TimerEnabled = True

    Catch ex As Exception

    End Try

  End Sub

  ''' <summary>
  ''' When a user clicks on any controls on one of your web pages, this function is then called with the post data. You can then parse the data and process as needed.
  ''' </summary>
  ''' <param name="page">The name of the page as registered with hs.RegisterLink or hs.RegisterConfigLink</param>
  ''' <param name="data">The post data</param>
  ''' <param name="user">The name of logged in user</param>
  ''' <param name="userRights">The rights of logged in user</param>
  ''' <returns>Any serialized data that needs to be passed back to the web page, generated by the clsPageBuilder class</returns>
  ''' <remarks></remarks>
  Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String

    Try

      WriteMessage("Entered postBackProc() Function.", MessageType.Debug)

      Dim postData As NameValueCollection = HttpUtility.ParseQueryString(data)

      '
      ' Write debug to console
      '
      If gLogLevel >= MessageType.Debug Then
        For Each keyName As String In postData.AllKeys
          Console.WriteLine(String.Format("{0}= {1}", keyName, postData(keyName)))
        Next
      End If

      '
      ' Process the post data
      '
      Select Case postData("id")
        Case "tabStatus"
          BuildTabStatus(True)

        Case "tabOptions"
          BuildTabOptions(True)

        Case "tabCameras"
          Me.pageCommands.Add("starttimer", "")

        Case "tabEvents"
          BuildTabEvents(True)

        Case "NetatmoUsername"
          Dim value As String = postData(postData("id"))
          SaveSetting("API", "Username", value)

          PostMessage("The Netatmo Username has been updated.")

        Case "NetatmoPassword"
          Dim value As String = postData(postData("id"))
          SaveSetting("API", "Password", value)

          PostMessage("The Netatmo Password has been updated.")

        Case "APIClientId"
          Dim value As String = postData(postData("id"))
          SaveSetting("API", "ClientId", value)

          gAPIClientId = value

          PostMessage("The Netatmo API Client Id has been updated.")

        Case "APIClientSecret"
          Dim value As String = postData(postData("id"))
          SaveSetting("API", "ClientSecret", value)

          gAPIClientSecret = value

          PostMessage("The Netatmo API Client Secret has been updated.")

        Case "selUnitType"
          Dim value As String = postData(postData("id"))
          SaveSetting("Options", "UnitType", value)

          PostMessage("The Unit Type has been updated.")

        Case "selVideoQuality"
          Dim value As String = postData(postData("id"))
          SaveSetting("Options", "VideoQuality", value)

          PostMessage("The Welcome Video Quality has been updated.")

        Case "selUpdateFrequency"
          Dim value As String = postData(postData("id"))
          SaveSetting("Options", "WelcomeUpdate", value)

          PostMessage("The Welcome Update frequency has been updated.")

        Case "selSnapshotsMaxWidth"
          Dim value As String = postData(postData("id"))
          SaveSetting("Options", "SnapshotsMaxWidth", value)

          Me.divToUpdate.Add("divCameras", BuildTabCameras())

        Case "selSnapshotEventMaxWidth"
          Dim value As String = postData(postData("id"))
          SaveSetting("Options", "SnapshotEventMaxWidth", value)

          Me.divToUpdate.Add("divEvents", BuildTabEvents())

        Case "selLogLevel"
          gLogLevel = Int32.Parse(postData("selLogLevel"))
          hs.SaveINISetting("Options", "LogLevel", gLogLevel.ToString, gINIFile)

          PostMessage("The application logging level has been updated.")

        Case "WebPageAccess_Guest"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Guest") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_GUEST
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_GUEST
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Normal"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Normal") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_NORMAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_NORMAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Local"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Local") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_LOCAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_LOCAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "timer" ' This stops the timer and clears the message
          If TimerEnabled Then 'this handles the initial timer post that occurs immediately upon enabling the timer.
            TimerEnabled = False
          Else
            Me.pageCommands.Add("stoptimer", "")
            Me.divToUpdate.Add("divMessage", "&nbsp;")
          End If

      End Select

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "postBackProc")
    End Try

    Return MyBase.postBackProc(page, data, user, userRights)

  End Function

#End Region

#Region "HSPI - Web Authorization"

  ''' <summary>
  ''' Returns the HTML Not Authorized web page
  ''' </summary>
  ''' <param name="LoggedInUser"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function WebUserNotUnauthorized(LoggedInUser As String) As String

    Try

      Dim sb As New StringBuilder

      sb.AppendLine("<table border='0' cellpadding='2' cellspacing='2' width='575px'>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td nowrap>")
      sb.AppendLine("     <h4>The Web Page You Were Trying To Access Is Restricted To Authorized Users ONLY</h4>")
      sb.AppendLine("   </td>")
      sb.AppendLine("  </tr>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td>")
      sb.AppendLine("     <p>This page is displayed if the credentials passed to the web server do not match the ")
      sb.AppendLine("      credentials required to access this web page.</p>")
      sb.AppendFormat("     <p>If you know the <b>{0}</b> user should have access,", LoggedInUser)
      sb.AppendFormat("      then ask your <b>HomeSeer Administrator</b> to check the <b>{0}</b> plug-in options", IFACE_NAME)
      sb.AppendFormat("      page to make sure the roles assigned to the <b>{0}</b> user allow access to this", LoggedInUser)
      sb.AppendLine("        web page.</p>")
      sb.AppendLine("  </td>")
      sb.AppendLine(" </tr>")
      sb.AppendLine(" </table>")

      Return sb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "WebUserNotUnauthorized")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

End Class