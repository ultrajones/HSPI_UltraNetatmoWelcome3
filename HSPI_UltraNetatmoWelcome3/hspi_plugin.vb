Imports System.Threading
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Net
Imports HomeSeerAPI
Imports Scheduler
Imports System.ComponentModel
Imports System.IO

Module hspi_plugin

  '
  ' Declare public objects, not required by HomeSeer
  '
  Dim actions As New hsCollection
  Dim triggers As New hsCollection
  Dim conditions As New Hashtable

  Const Pagename = "Events"

  Public HSDevices As New SortedList

  Public NetatmoAPI As hspi_netatmo_api
  Public gAPIClientId As String = String.Empty
  Public gAPIClientSecret As String = String.Empty
  Public gAPIUsername As String = String.Empty
  Public gAPIPassword As String = String.Empty
  Public gAPIScope As String = "read_camera access_camera write_camera read_presence access_presence"

  Public gAPICultureInfo As String = "en-us"

  Public gHomesId As String = String.Empty
  Public gWelcomeUpdate As String = "1"

  Public gNetatmoPerson As New List(Of hspi_netatmo_api.Person)
  Public gNetatmoPersonLock As New Object

  Public gNetatmoEvents As New List(Of hspi_netatmo_api.Event)
  Public gNetatmoEventsLock As New Object
  Public gNetatmoEventCount As UInteger = 0

  Public gNetatmoCameras As New List(Of hspi_netatmo_api.Camera)
  Public gNetatmoCamerasLock As New Object

  Public Const IFACE_NAME As String = "UltraNetatmoWelcome3"

  Public Const LINK_TARGET As String = "hspi_ultranetatmowelcome3/hspi_ultranetatmowelcome3.aspx"
  Public Const LINK_URL As String = "hspi_ultranetatmowelcome3.html"
  Public Const LINK_TEXT As String = "UltraNetatmoWelcome3"
  Public Const LINK_PAGE_TITLE As String = "UltraNetatmoWelcome3 HSPI"
  Public Const LINK_HELP As String = "/hspi_ultranetatmowelcome3/UltraNetatmoWelcome3_HSPI_Users_Guide.pdf"

  Public gBaseCode As String = String.Empty
  Public gIOEnabled As Boolean = True
  Public gImageDir As String = "/images/hspi_ultranetatmowelcome3/"
  Public gHSInitialized As Boolean = False
  Public gINIFile As String = "hspi_" & IFACE_NAME.ToLower & ".ini"

  Public gRTCReceived As Boolean = False
  Public gDeviceValueType As String = "1"
  Public gDeviceImage As Boolean = True
  Public gStatusImageSizeWidth As Integer = 32
  Public gStatusImageSizeHeight As Integer = 32

  Public gEventArchiveToDir As Boolean = False            ' Indicates if events should be archived to archive directory
  Public gEventArchiveToFTP As Boolean = False            ' Indicates if events should be archived to FTP directory
  Public gEventEmailNotification As Boolean = False       ' Indicates if events should generate an e-mail notification
  Public gEventCompressToZip As Boolean = True            ' Indicates if event snapshots should be compressed
  Public gEventSnapshotMax As Integer = 50                ' The default number of events to store per NetCam
  Public gEventVideoQuality As String = "disabled"        ' The qualitify of the downloaded event

  Public gSnapshotMaxWidth As String = "Auto"             ' The size of the snapshot image
  Public gSnapshotRefreshInterval As Integer = 0          ' The default snapshot refresh interval
  Public gSnapshotEventMaxWidth As String = "160px"       ' The size of the snapshot event images

  Public gMonitoring As Boolean = True

  Public gHSAppPath As String = ""
  Public gImagePersonDirectory As String = ""
  Public gImageCameraDirectory As String = ""
  Public gImageEventDirectory As String = ""

#Region "UltraNetatmoWelcome3 Public Functions"

  ''' <summary>
  ''' Welcome Event Thread
  ''' </summary>
  ''' <remarks></remarks>
  Friend Sub CheckCameraThread()

    Dim strMessage As String = ""
    Dim iCheckInterval As Integer = 0

    Dim bAbortThread As Boolean = False

    Try
      WriteMessage("The CheckCameraThread thread has started ...", MessageType.Debug)

      '
      ' Begin Main While Loop
      '
      While bAbortThread = False

        If gMonitoring = True Then

          Try

            Dim unitType As String = hs.GetINISetting("Options", "UnitType", "0", gINIFile)

            Dim HomeData As hspi_netatmo_api.HomeData = NetatmoAPI.GetHomeData
            If Not IsNothing(HomeData.status) AndAlso HomeData.status.Length > 0 Then

              For Each Home As hspi_netatmo_api.Home In HomeData.body.homes
                '
                ' Process Welcome Homes Data
                '
                If Not IsNothing(Home) Then
                  Dim homeId As String = Home.id
                  Dim homeName As String = Home.name
                  Dim homeCity As String = Home.place.city
                  Dim homeCountry As String = Home.place.country
                  Dim homeTimeZone As String = Home.place.timezone

                  If homeId.Length > 0 And gHomesId <> homeId Then
                    gHomesId = homeId
                  End If

                  '
                  ' Process Each Welcome Person
                  '
                  SyncLock gNetatmoPersonLock

                    '
                    ' Array of persons seen in the Home. Array structure: last_seen, face, out_of_sight, pseudo.
                    '   If "pseudo" Is missing Then from the array, the person Is unknown
                    '
                    Dim iPersons As Integer = 0
                    For Each Person As hspi_netatmo_api.Person In Home.persons
                      Dim psudeo As String = Person.pseudo

                      iPersons += 1

                      If psudeo.Length > 0 Then
                        Dim last_seen As String = Person.last_seen
                        Dim out_of_site As String = Person.out_of_sight

                        Dim dv_root_addr As String = "Person-Root"
                        Dim dv_root_type As String = "Welcome Person"
                        Dim dv_root_name As String = "Person Root Device"

                        Dim dv_addr As String = String.Format("Person-{0}", psudeo)
                        Dim dv_name As String = psudeo
                        Dim dv_type As String = "Welcome Person"

                        Dim HomeAway As String = IIf(out_of_site = True, "Away", "Home")
                        Dim dv_value As Integer = IIf(out_of_site = True, 0, 1)

                        '
                        ' Update the Welcome Person HomeSeer device
                        '
                        GetHomeSeerDevice(dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr, Person.id)
                        SetDeviceValue(dv_addr, dv_value)

                        Try
                          Dim startDate As Long = Val(last_seen)
                          Dim timeDiff As String = GetTimeDiff(startDate, DateTime.Now)
                          Dim dv_string As String = String.Format("{0} [{1}]", HomeAway, timeDiff)

                          SetDeviceString(dv_addr, dv_string)

                          If gNetatmoPerson.Any(Function(s) s.id = Person.id) = False Then
                            '
                            ' This is a new person
                            '
                            Dim Face As hspi_netatmo_api.face = Person.face
                            Dim strSnapshotFilename As String = FixPath(String.Format("{0}/{1}.jpg", gImagePersonDirectory, dv_addr.ToLower))
                            If File.Exists(strSnapshotFilename) = False Then
                              Dim WebExceptionStatus As WebExceptionStatus = NetatmoAPI.GetCameraPicture(Face.id, Face.key, strSnapshotFilename, String.Empty, 30)

                            End If

                            gNetatmoPerson.Add(Person)
                          End If

                        Catch pEx As Exception
                          Call ProcessError(pEx, "CheckCameraThread()")
                        End Try

                      Else
                        '
                        ' Unknown person detected
                        '
                      End If
                    Next

                  End SyncLock

                  '
                  ' Array of all cameras in the Home. Array structure: status, sd_status, alim_status, is_locale, vpn_url.
                  '
                  SyncLock gNetatmoCamerasLock

                    For Each Camera As hspi_netatmo_api.Camera In Home.cameras
                      Dim camera_id As String = Camera.id
                      Dim camera_name As String = Camera.name
                      Dim camera_status As String = Camera.status
                      Dim sd_status As String = Camera.sd_status
                      Dim alim_status As String = Camera.alim_status
                      Dim is_local As Boolean = Camera.is_local

                      If gNetatmoCameras.Any(Function(s) s.id = Camera.id) = False Then
                        '
                        ' This is a new camera
                        '
                        gNetatmoCameras.Add(Camera)
                      End If

                      If camera_id.Length > 0 And camera_name.Length > 0 Then
                        Dim dv_root_addr As String = "Camera-Root"
                        Dim dv_root_type As String = "Welcome Camera"
                        Dim dv_root_name As String = "Camera Root Device"

                        Dim dv_addr As String = String.Format("Camera-{0}", camera_id.Replace(":", "-"))
                        Dim dv_name As String = camera_name
                        Dim dv_type As String = "Welcome Camera"

                        '
                        ' Determine Status
                        '
                        Dim dv_value As Integer = IIf(camera_status.ToLower = "on", CameraStatus.Online, CameraStatus.Offline)
                        If sd_status.ToLower <> "on" Then
                          dv_value = CameraStatus.SDCard
                        ElseIf alim_status.ToLower <> "on" Then
                          dv_value = CameraStatus.PowerSupply
                        ElseIf is_local <> True Then
                          If dv_value = CameraStatus.Online Then
                            dv_value = CameraStatus.NotLocal
                          End If
                        End If

                        '
                        ' Update the Welcome Camera HomeSeer device
                        '
                        GetHomeSeerDevice(dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr, Camera.id)
                        SetDeviceValue(dv_addr, dv_value)

                        '
                        ' Get a snapshot from the camera
                        '
                        If dv_value = CameraStatus.Online Then
                          'If dv_value = CameraStatus.Online OrElse dv_value = CameraStatus.NotLocal Then

                          Camera.camera_url = Camera.vpn_url
                          If is_local = True Then
                            Dim WelcomePing As hspi_netatmo_api.WelcomePing = NetatmoAPI.SendWelcomePing(Camera.vpn_url)
                            If WelcomePing.local_url.Length > 0 Then
                              Camera.camera_url = WelcomePing.local_url
                            End If
                          End If
                          Dim strSnapshotFilename As String = FixPath(String.Format("{0}/{1}.jpg", gImageCameraDirectory, dv_addr.ToLower))
                          Dim WebExceptionStatus As WebExceptionStatus = NetatmoAPI.GetLiveSnapshot(Camera.camera_url, strSnapshotFilename, String.Empty, 30)

                          If WebExceptionStatus = WebExceptionStatus.Success Then
                            If File.Exists(strSnapshotFilename) = True Then
                              File.SetCreationTime(strSnapshotFilename, DateTime.Now)
                            End If
                          End If
                        End If

                      End If

                    Next
                  End SyncLock

                  '
                  ' Array of events occured in the Home.  
                  '  Warning:  Welcome And Presence events don't have the same structure. See each products page for more details.
                  '
                  SyncLock gNetatmoEventsLock

                    For Each [Event] As hspi_netatmo_api.Event In Home.events

                      If [Event].video_status <> "recording" Then

                        If gNetatmoEvents.Any(Function(s) s.id = [Event].id) = False Then
                          '
                          ' This event is new
                          '
                          gNetatmoEvents.Add([Event])

                          '
                          ' Populate the Welcome Camera variables
                          '
                          Dim camera_name As String = "Unknown"
                          Dim camera_url As String = String.Empty
                          Dim camera_vpn_url As String = String.Empty
                          Dim camera_is_local As Boolean = False

                          '
                          ' Get the welcome camera data based on event camera_id
                          '
                          SyncLock gNetatmoCamerasLock
                            If gNetatmoCameras.Any(Function(s) s.id = [Event].camera_id) = True Then
                              Dim NetatmoCamera As hspi_netatmo_api.Camera = gNetatmoCameras.Find(Function(s) s.id = [Event].camera_id)
                              camera_name = NetatmoCamera.name
                              camera_url = NetatmoCamera.camera_url
                              camera_vpn_url = NetatmoCamera.vpn_url
                              camera_is_local = NetatmoCamera.is_local
                            End If
                          End SyncLock

                          '
                          ' Populate Welcome Event Variables
                          '
                          Dim eventMessage As String = [Event].message
                          Dim eventSubMessage As String = [Event].sub_message
                          Dim eventVideo As String = [Event].video_status.ToLower
                          Dim eventType As String = [Event].type.ToLower
                          Dim eventTime As DateTime = ConvertEpochToDateTime([Event].time)

                          '
                          ' Process Welcome Event Type
                          '
                          Select Case eventType
                            Case "alarm_started"
                              '
                              ' Process Alarm Condition
                              '
                              If gNetatmoEventCount > 0 Then
                                WriteMessage(String.Format("{0} reported by {1} at {2}", eventMessage, camera_name, eventTime.ToShortTimeString), MessageType.Informational)
                              End If

                            Case "sd"
                              '
                              ' Event triggered by the SD card status change
                              '

                            Case "alim"
                              '
                              ' Event triggered by the power supply status change
                              '

                            Case "boot"
                              '
                              ' Process Boot Event (no video) - When the camera Is booting
                              '
                              If gNetatmoEventCount > 0 Then
                                WriteMessage(String.Format("{0} reported by {1} at {2}", eventMessage, camera_name, eventTime.ToShortTimeString), MessageType.Informational)

                                Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.WelcomeCameraBoot, [Event].camera_id)
                                hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.WelcomeCameraBoot, strTrigger)
                              End If

                            ' {
                            '"id":"5652769749c75fb746b35146",
                            '"type":"boot",
                            '"time":1448244886,
                            '"camera_id":"70:ee:50:13:93:86",
                            '"message":"Den started"
                            '}

                            Case "on"
                              '
                              ' Process the welcome Monitoring event (no video) - Whenever monitoring is activated
                              '
                              If gNetatmoEventCount > 0 Then
                                WriteMessage(String.Format("{0} reported by {1} at {2}", eventMessage, camera_name, eventTime.ToShortTimeString), MessageType.Informational)

                                Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.MonitoringEnabled, [Event].camera_id)
                                hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.MonitoringEnabled, strTrigger)
                              End If

                            Case "off"
                              '
                              ' Process the welcome Monitoring event (no video) - Whenever monitoring is suspended
                              '
                              If gNetatmoEventCount > 0 Then
                                WriteMessage(String.Format("{0} reported by {1} at {2}", eventMessage, camera_name, eventTime.ToShortTimeString), MessageType.Informational)

                                Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.MonitoringDisabled, [Event].camera_id)
                                hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.MonitoringDisabled, strTrigger)
                              End If

                            Case "connection"
                              '
                              ' Process the welcome camera connection event (no video) - When the camera connects to Netatmo servers
                              '
                              If gNetatmoEventCount > 0 Then
                                WriteMessage(String.Format("{0} reported by {1} at {2}", eventMessage, camera_name, eventTime.ToShortTimeString), MessageType.Informational)

                                Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.CameraConnected, [Event].camera_id)
                                hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.CameraConnected, strTrigger)
                              End If

                            Case "disconnection"
                              '
                              ' Process the welcome camera connection event (no video) - When the camera loses connection to Netatmo servers
                              '
                              If gNetatmoEventCount > 0 Then
                                WriteMessage(String.Format("{0} reported by {1} at {2}", eventMessage, camera_name, eventTime.ToShortTimeString), MessageType.Informational)

                                Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.CameraDisconnected, [Event].camera_id)
                                hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.CameraDisconnected, strTrigger)
                              End If

                            '{
                            '"id":"565265c10ce67ca074c728d1",
                            '"type":"connection",
                            '"time":1448240577,
                            '"camera_id":"70:ee:50:13:a7:12",
                            '"message":"Family Room connected"
                            '}

                            Case "person"
                              '
                              ' Process the welcome person event (video) - Event triggered when Welcome detects a face
                              '
                              If gNetatmoEventCount > 0 Then
                                WriteMessage(String.Format("{0} reported by {1} at {2}", eventMessage, camera_name, eventTime.ToShortTimeString), MessageType.Informational)

                                If Regex.IsMatch(eventMessage, "Unknown.*seen") = True Then
                                  Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.SomeoneUnknownSceen, [Event].camera_id)
                                  hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.SomeoneUnknownSceen, strTrigger)

                                ElseIf Regex.IsMatch(eventMessage, "seen") = True Then

                                  If [Event].is_arrival = True Then
                                    Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.SomeoneKnownArrives, [Event].camera_id)
                                    hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.SomeoneKnownArrives, strTrigger)
                                  Else
                                    Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.SomeoneknownSceen, [Event].camera_id)
                                    hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.SomeoneknownSceen, strTrigger)
                                  End If

                                End If
                              End If

                            Case "person_away"
                              '
                              ' Process the welcome person_away event (no video) - Event triggered when geofencing implies the person has left the home
                              '
                              If gNetatmoEventCount > 0 Then
                                WriteMessage(String.Format("{0} reported by {1} at {2}", eventMessage, camera_name, eventTime.ToShortTimeString), MessageType.Informational)

                                Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.SomeoneKnownLeaves, [Event].camera_id)
                                hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.SomeoneKnownLeaves, strTrigger)
                              End If

                              '{
                              '"id":"565356a265d1c42b0c34afd7",
                              '"type":"person_away",
                              '"time":1448302239,
                              '"camera_id":"70:ee:50:13:a7:12",
                              '"person_id":"b07d88c3-bb16-4bc3-890a-a6889828e993",
                              '"message":"<b>Randy<\/b> left",
                              '"sub_message":"After not being seen for a while, people are listed in this category. <b>The delay is customizable in the Settings<\/b> (section \u00ab Adjust Welcome to your needs \u00bb)"
                              '}

                            Case "person_home"

                              If gNetatmoEventCount > 0 Then
                                WriteMessage(String.Format("{0} reported by {1} at {2}", eventMessage, camera_name, eventTime.ToShortTimeString), MessageType.Informational)

                                Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.SomeoneKnownArrives, [Event].camera_id)
                                hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.SomeoneKnownArrives, strTrigger)
                              End If

                              '{
                              '"id":"5d1f5e08ad7ab3d6e87a62a8",
                              '"type":"person_away",
                              '"time":1562336776,
                              '"camera_id":"70:ee:50:13:a7:12",
                              '"device_id":"70:ee:50:13:a7:12",
                              '"person_id":"3daeede7-dd89-455a-897c-0b8b23a8e87d",
                              '"message":"ULTRANETATMOWELCOME3 HSPI set Kyle as \"Away\"",
                              '"sub_message":"ULTRANETATMOWELCOME3 HSPI told Welcome to set Kyle as \"Away\"."
                              '},

                            Case "movement"
                              '
                              ' Process the welcome person event (video) - Event triggered when Welcome detects a motion
                              '
                              If gNetatmoEventCount > 0 Then
                                WriteMessage(String.Format("{0} reported by {1} at {2}", eventMessage, camera_name, eventTime.ToShortTimeString), MessageType.Informational)

                                Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.MotionDetected, [Event].camera_id)
                                hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.MotionDetected, strTrigger)
                              End If

                              '{
                              '"id":"56525f1e65d1c4c798b21ea5",
                              '"type":"person",
                              '"time":1448238875,
                              '"camera_id":"70:ee:50:13:a7:12",
                              '"person_id":"ec10a1e3-2894-4041-889f-f25fae691e69",
                              '"snapshot":{
                              '"id":"56525f1e65d1c4c798b21ea6",
                              '"version":1,
                              '"key":"1deccbd6ec8a3403e79a92b23c0ac2c41ec38bd8b319461b1d2781032b47e769"
                              '},
                              '"video_id":"74d46c86-113e-452a-8ece-785deb4fe4fe",
                              '"video_status":"available",
                              '"is_arrival":false,
                              '"message":"<b>Rachel<\/b> seen"
                              '}

                            Case "outdoor"
                              '
                              ' Event triggered when Presence detects a human, a car or an animal
                              '
                              If gNetatmoEventCount > 0 Then

                                For Each Outdoor_Event As hspi_netatmo_api.Outdoor_Event In [Event].event_list

                                  Dim outdoorEventType As String = Outdoor_Event.type.ToLower
                                  Dim outdoorEventTime As DateTime = ConvertEpochToDateTime(Outdoor_Event.time)
                                  Dim outdoorEventMessage As String = Outdoor_Event.message

                                  WriteMessage(String.Format("{0} reported by {1} at {2}", outdoorEventMessage, camera_name, outdoorEventTime.ToShortTimeString), MessageType.Informational)

                                  Select Case outdoorEventType
                                    Case "human"
                                      Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.OutdoorEventHumanSeen, [Event].camera_id)
                                      hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.OutdoorEventHumanSeen, strTrigger)

                                    Case "vehicle"
                                      Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.OutdoorEventVehicleSeen, [Event].camera_id)
                                      hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.OutdoorEventVehicleSeen, strTrigger)

                                    Case "animal"
                                      Dim strTrigger As String = String.Format("{0},{1}", WelcomeTriggers.OutdoorEventAnimalSeen, [Event].camera_id)
                                      hspi_plugin.CheckTrigger(IFACE_NAME, NetatmoTriggers.WelcomeTriggers, WelcomeTriggers.OutdoorEventAnimalSeen, strTrigger)

                                  End Select

                                Next

                              End If

                            Case "daily_summary"
                              '
                              ' Event triggered when the video summary of the last 24 hours is available
                              '

                            Case "new_module"
                              '
                              ' A new module has been paired with Welcome
                              '

                            Case "module_connect"
                              '
                              ' Module is connected with Welcome (after disconnection)
                              '

                            Case "module_disconnect"
                              '
                              ' Module lost its connection with Welcome
                              '

                            Case "module_low_battery"
                              '
                              ' Module's battery is low
                              '

                            Case "module_end_update"
                              '
                              ' Module's firmware update is over
                              '

                            Case "tag_big_move"
                              '
                              ' Tag detected a big move
                              '

                            Case "tag_small_move"
                              '
                              ' Tag detected a small move
                              '

                            Case "tag_uninstalled"
                              '
                              ' Tag was uninstalled
                              '

                            Case "tag_open"
                              '
                              ' Tag detected the door/window was left open
                              '

                            Case "qrcode_detected"
                              '
                              ' qrcode detected
                              '

                            Case Else
                              '
                              ' Process unknown event type
                              '
                              WriteMessage(String.Format("Unhandled event type {0}", eventType), MessageType.Warning)
                          End Select
                        End If
                      End If

                    Next

                    '
                    ' Update our Event Count
                    '
                    gNetatmoEventCount = gNetatmoEvents.Count

                  End SyncLock

                End If

              Next

            End If

          Catch pEx As Exception
            '
            ' Process the error
            '
            Call ProcessError(pEx, "CheckCameraThread()")
          End Try

        End If

        '
        ' Sleep the requested number of minutes between runs
        '
        iCheckInterval = CInt(hs.GetINISetting("Options", "WelcomeUpdate", gWelcomeUpdate, gINIFile))
        Thread.Sleep(1000 * (60 * iCheckInterval))

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("CheckCameraThread received abort request, terminating normally."), MessageType.Debug)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "CheckCameraThread()")
    End Try

  End Sub

  ''' <summary>
  ''' Welcome Event Thread
  ''' </summary>
  ''' <remarks></remarks>
  Friend Sub CheckCameraEventThread()

    Dim strMessage As String = ""
    Dim iCheckInterval As Integer = 0

    Dim bAbortThread As Boolean = False

    Try
      WriteMessage("The CheckCameraEventThread thread has started ...", MessageType.Debug)

      Thread.Sleep(1000 * 10)

      '
      ' Begin Main While Loop
      '
      While bAbortThread = False

        Try

          SyncLock gNetatmoEventsLock
            '
            ' Download Events
            '
            For Each [Event] As hspi_netatmo_api.Event In gNetatmoEvents.FindAll(Function(s) s.video_status = "available" And s.video_downloaded = False And s.download_tries <= 4)
              Dim eventVideo As String = [Event].video_status.ToLower
              Dim eventType As String = [Event].type.ToLower
              Dim eventTime As DateTime = ConvertEpochToDateTime([Event].time)

              Dim camera_name As String = "Unknown"
              Dim camera_status As String = String.Empty
              Dim camera_url As String = String.Empty
              Dim camera_vpn_url As String = String.Empty
              Dim camera_is_local As Boolean = False

              '
              ' Increment the tries
              '
              [Event].download_tries += 1
              If [Event].download_tries > 4 Then
                If gEventVideoQuality <> "disabled" Then
                  WriteMessage(String.Format("{0} {1}", [Event].id, eventTime.ToString), MessageType.Error)
                End If
              End If

              '
              ' Get the welcome camera data based on event camera_id
              '
              SyncLock gNetatmoCamerasLock
                If gNetatmoCameras.Any(Function(s) s.id = [Event].camera_id) = True Then
                  Dim NetatmoCamera As hspi_netatmo_api.Camera = gNetatmoCameras.Find(Function(s) s.id = [Event].camera_id)
                  camera_name = NetatmoCamera.name
                  camera_status = NetatmoCamera.status
                  camera_url = NetatmoCamera.camera_url
                  camera_vpn_url = NetatmoCamera.vpn_url
                  camera_is_local = NetatmoCamera.is_local
                End If
              End SyncLock

              '
              ' Determine if the camera is local
              '
              If camera_is_local = False Then
                camera_url = camera_vpn_url
              End If

              '
              ' Prepare the Event Directory
              '
              Dim EventDirectory As String = FixPath(String.Format("{0}/{1}", gImageEventDirectory, [Event].id.ToLower))
              If Directory.Exists(EventDirectory) = False Then
                Directory.CreateDirectory(EventDirectory)
                Directory.SetCreationTime(EventDirectory, eventTime)
              End If

              '
              ' Get the Event Snapshots
              '
              If [Event].type = "outdoor" Then

                Dim snapshotFileAvailable As Integer = [Event].event_list.Count
                Dim snapshotFileDownloaded As Integer = 0

                For Each Outdoor_Event As hspi_netatmo_api.Outdoor_Event In [Event].event_list

                  If Outdoor_Event.snapshot.filename.Length > 0 Then

                    '
                    ' Get a snapshot from the camera
                    '
                    If camera_status = "on" Then

                      Dim strSnapshotFilename As String = FixPath(String.Format("{0}/{1}.jpg", EventDirectory, Outdoor_Event.id.ToLower))
                      Dim WebExceptionStatus As WebExceptionStatus = NetatmoAPI.GetEventSnapshot(camera_url, Outdoor_Event.snapshot.filename, strSnapshotFilename, String.Empty, 30)

                      If WebExceptionStatus = WebExceptionStatus.Success Then
                        If File.Exists(strSnapshotFilename) = True Then
                          File.SetCreationTime(strSnapshotFilename, DateTime.Now)
                        End If
                      End If

                      If File.Exists(strSnapshotFilename) = True Then
                        snapshotFileDownloaded += 1
                      End If

                    End If

                  Else

                    Dim strSnapshotFilename As String = FixPath(String.Format("{0}/{1}.jpg", EventDirectory, Outdoor_Event.snapshot.id.ToLower))
                    If File.Exists(strSnapshotFilename) = False Then
                      Dim WebExceptionStatus As WebExceptionStatus = NetatmoAPI.GetCameraPicture(Outdoor_Event.snapshot.id, Outdoor_Event.snapshot.key, strSnapshotFilename, String.Empty, 30)

                      If WebExceptionStatus = WebExceptionStatus.Success Then
                        If File.Exists(strSnapshotFilename) = True Then
                          File.SetCreationTime(strSnapshotFilename, eventTime)
                        End If
                      End If
                    End If

                    If File.Exists(strSnapshotFilename) = True Then
                      snapshotFileDownloaded += 1
                    End If

                  End If

                Next

                If snapshotFileAvailable = snapshotFileDownloaded Then
                  '
                  ' Indicate the Video has been downloaded
                  '
                  [Event].snapshot_downloaded = True
                End If

              Else

                Dim strSnapshotFilename As String = FixPath(String.Format("{0}/{1}.jpg", EventDirectory, [Event].snapshot.id.ToLower))
                If File.Exists(strSnapshotFilename) = False Then
                  Dim WebExceptionStatus As WebExceptionStatus = NetatmoAPI.GetCameraPicture([Event].snapshot.id, [Event].snapshot.key, strSnapshotFilename, String.Empty, 30)

                  If WebExceptionStatus = WebExceptionStatus.Success Then
                    If File.Exists(strSnapshotFilename) = True Then
                      File.SetCreationTime(strSnapshotFilename, eventTime)
                    End If
                  End If
                Else
                  [Event].snapshot_downloaded = True
                End If

              End If

              If gEventVideoQuality = "disabled" Then
                [Event].video_downloaded = True
              End If

              '
              ' Check to see if the camera is connected
              '
              If camera_url.Length > 0 AndAlso camera_status = "on" AndAlso [Event].video_downloaded = False Then
                '
                ' Get the Event Video [video_id].m3u8
                '
                Dim eventIdM3u8 As String = FixPath(String.Format("{0}/{1}/{1}.m3u8", gImageEventDirectory, [Event].id.ToLower))
                If File.Exists(eventIdM3u8) = False Then

                  Dim video_url As String = String.Format("vod/{0}/{1}.m3u8", [Event].video_id, "index")
                  Dim address As String = String.Format("{0}/{1}", camera_url, video_url)

                  Try

                    If Directory.Exists(EventDirectory) = True Then
                      My.Computer.Network.DownloadFile(address, eventIdM3u8)
                    End If

                  Catch pEx As WebException
                    WriteMessage(String.Format("Unable To download video event Id {0} [{1}].  {2}", [Event].id, address, pEx.Message), MessageType.Warning)

                  Catch pEx As Exception
                    Call ProcessError(pEx, "CheckCameraEventThread()")
                  End Try

                End If

                '
                ' Get the Event Video index.m3u8
                '
                Dim videoIndexm3u8 As String = FixPath(String.Format("{0}/{1}/{2}", gImageEventDirectory, [Event].id.ToLower, "index.m3u8"))

                If File.Exists(eventIdM3u8) = True Then
                  '
                  ' Adjust the Video URL depending on content of m3u8 file
                  '
                  Dim video_filename As String = String.Format("files/{0}/index.m3u8", gEventVideoQuality)
                  Dim video_url As String = String.Format("vod/{0}/{1}", [Event].video_id, video_filename)

                  Using readStream As New StreamReader(eventIdM3u8, Encoding.UTF8)
                    Do While Not readStream.EndOfStream
                      Dim strLine As String = readStream.ReadLine

                      Dim regexPattern As String = String.Format("files\/{0}\/index", gEventVideoQuality)
                      If Regex.IsMatch(strLine, regexPattern) = True Then
                        video_url = String.Format("vod/{0}/{1}", [Event].video_id, strLine)
                        Exit Do
                      End If
                    Loop
                  End Using

                  If File.Exists(videoIndexm3u8) = False Then

                    Dim address As String = String.Format("{0}/{1}", camera_url, video_url)

                    Try

                      If Directory.Exists(EventDirectory) = True Then
                        My.Computer.Network.DownloadFile(address, videoIndexm3u8)
                      End If

                    Catch pEx As WebException
                      WriteMessage(String.Format("Unable To download video event Id {0} [{1}].  {2}", [Event].id, address, pEx.Message), MessageType.Warning)

                    Catch pEx As Exception
                      Call ProcessError(pEx, "CheckCameraEventThread()")
                    End Try

                  End If

                End If

                '
                ' Get the Individual Video Files
                '
                If File.Exists(videoIndexm3u8) = True Then
                  Dim videoFileAvailable As Integer = 0
                  Dim videoFileDownloaded As Integer = 0

                  '
                  ' Open the file using a stream reader
                  '
                  Using readStream As New StreamReader(videoIndexm3u8, Encoding.UTF8)

                    Do While Not readStream.EndOfStream
                      Dim strLine As String = readStream.ReadLine

                      ' http://192.168.2.38/f49d998d3534df4e2883e64de70b8d04/vod/b5f31fa5-140f-428d-8479-c6d4edd4a876/files/poor/live0001181626.ts
                      If Regex.IsMatch(strLine, "live\d+\.ts$") Then
                        Dim strFileName As String = Regex.Match(strLine, "(live\d+\.ts)$").ToString
                        Dim videoFileName As String = FixPath(String.Format("{0}/{1}/{2}", gImageEventDirectory, [Event].id.ToLower, strFileName))

                        If File.Exists(videoFileName) = False Then
                          Dim video_filename As String = String.Format("files/{0}", gEventVideoQuality)
                          Dim video_url As String = String.Format("/vod/{0}/{1}/{2}", [Event].video_id, video_filename, strLine)
                          Dim address As String = String.Format("{0}{1}", camera_url, video_url)

                          If strLine.StartsWith("http") = True Then
                            address = strLine
                          End If

                          Try

                            videoFileAvailable += 1

                            If Directory.Exists(EventDirectory) = True Then
                              My.Computer.Network.DownloadFile(address, videoFileName)
                            End If

                            videoFileDownloaded += 1

                          Catch pEx As WebException
                            WriteMessage(String.Format("Unable To download video event Id {0} [{1}].  {2}", [Event].id, address, pEx.Message), MessageType.Warning)

                          Catch pEx As Exception
                            Call ProcessError(pEx, "CheckCameraEventThread()")
                          End Try
                        End If

                      End If

                    Loop

                  End Using

                  If videoFileAvailable = videoFileDownloaded Then
                    '
                    ' Indicate the Video has been downloaded
                    '
                    [Event].video_downloaded = True
                  End If

                End If

                '
                ' Write HTML Index file
                '
                Dim indexFileName As String = FixPath(String.Format("{0}/{1}/{2}", gImageEventDirectory, [Event].id.ToLower, "index.html"))
                If File.Exists(indexFileName) = False Then
                  Using indexFile As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(indexFileName, False)

                    indexFile.WriteLine("<html lang=""en"">")
                    indexFile.WriteLine(" <head>")
                    indexFile.WriteLine("  <meta charset=utf-8/>")
                    indexFile.WriteLine(" </head>")
                    indexFile.WriteLine(" <body>")
                    indexFile.WriteLine("  <video controls height=""auto"" width=""100%"">")
                    indexFile.WriteLine("   <source src=""index.m3u8"" type=""application/x-mpegURL"">")
                    indexFile.WriteLine("  </video>")
                    indexFile.WriteLine(" </body>")
                    indexFile.WriteLine("</html>")
                    indexFile.Close()

                  End Using

                End If

                '
                ' Lets just download 1 event per round
                '
                Exit For

              End If

            Next

          End SyncLock

        Catch pEx As Exception
          '
          ' Process the error
          '
          Call ProcessError(pEx, "CheckCameraEventThread()")
        End Try

        '
        ' Sleep the requested number of minutes between runs
        '
        Thread.Sleep(1000)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("CheckCameraEventThread received abort request, terminating normally."), MessageType.Debug)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "CheckCameraEventThread()")
    End Try

  End Sub

  ''' <summary>
  ''' Performs Snapshot File Maintenance
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub CheckMaintenanceThread()

    Dim bAbortThread As Boolean = False

    Try
      WriteMessage("The CheckMaintenanceThread thread has started ...", MessageType.Debug)

      Thread.Sleep(1000 * 10)

      While bAbortThread = False

        Try
          '
          ' Purge Events
          '
          PurgeWelcomeEvents()

        Catch pEx As Exception
          '
          ' Return message
          '
          ProcessError(pEx, "CheckMaintenanceThread()")
        End Try

        '
        ' Give up some time
        '
        Thread.Sleep(1000 * 60)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("CheckMaintenanceThread thread received abort request, terminating normally."), MessageType.Informational)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "CheckMaintenanceThread()")
    Finally
      '
      ' Notify that we are exiting the thread
      '
      WriteMessage(String.Format("CheckMaintenanceThread terminated."), MessageType.Debug)
    End Try

  End Sub

  Private Function sortEvent(ByVal x As hspi_netatmo_api.Event, ByVal y As hspi_netatmo_api.Event) As Integer
    If (x.time < y.time) Then
      Return 1
    End If

    If (x.time > y.time) Then
      Return -1
    Else
      Return 0
    End If
  End Function

  ''' <summary>
  ''' Purges Welcome Events
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub PurgeWelcomeEvents()

    Try

      Dim LastDate As Date = Now
      Dim LastEvent As String = ""
      Dim iEventCount As Long = 0

      If gNetatmoEvents.Count > gEventSnapshotMax + 10 Then

        SyncLock gNetatmoEventsLock
          gNetatmoEvents.Sort(AddressOf sortEvent)

          Dim [Event] As hspi_netatmo_api.Event = gNetatmoEvents(gNetatmoEvents.Count - 1)
          Dim eventDate As DateTime = ConvertEpochToDateTime([Event].time)
          Dim event_id As String = [Event].id

          Dim bResult As Boolean = gNetatmoEvents.Remove([Event])
          If bResult = True Then
            WriteMessage(String.Format("Purging event {0} from {1}.", event_id, eventDate.ToString), MessageType.Debug)
          Else
            WriteMessage(String.Format("Purging event {0} from {1} failed.", event_id, eventDate.ToString), MessageType.Error)
          End If
        End SyncLock

        Dim EventDirectory As String = FixPath(gImageEventDirectory)
        Dim RootDirectoryInfo As New IO.DirectoryInfo(EventDirectory)
        For Each EventDirectoryInfo As DirectoryInfo In RootDirectoryInfo.GetDirectories

          Dim event_id As String = Regex.Match(EventDirectoryInfo.Name, "([a-z0-9]{24})$").ToString
          If gNetatmoEvents.Any(Function(s) s.id = event_id) = False Then

            WriteMessage(String.Format("Purging event {0} directory.", event_id), MessageType.Debug)
            Directory.Delete(EventDirectoryInfo.FullName, True)

          End If

        Next

      End If

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "PurgeWelcomeEvents()")
    End Try

  End Sub

  ''' <summary>
  ''' Gets plug-in setting from INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValueDefault"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSetting(ByVal strSection As String,
                             ByVal strKey As String,
                             ByVal strValueDefault As String) As String

    Dim strMessage As String = ""

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered GetSetting() Function.", MessageType.Debug)

      '
      ' Get the ini settings
      '
      Dim strValue As String = hs.GetINISetting(strSection, strKey, strValueDefault, gINIFile)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Check to see if we need to decrypt the data
      '
      If strSection = "API" And strKey = "Password" Then
        strValue = hs.DecryptString(strValue, "&Cul8r#1")
      End If

      Return strValue

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetSetting()")
      Return ""
    End Try

  End Function

  ''' <summary>
  '''  Saves plug-in settings to INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValue"></param>
  ''' <remarks></remarks>
  Public Sub SaveSetting(ByVal strSection As String,
                         ByVal strKey As String,
                         ByVal strValue As String)

    Dim strMessage As String = ""

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered SaveSetting() subroutine.", MessageType.Debug)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Apply the API Consumer Key
      '
      If strSection = "API" And strKey = "ClientId" Then
        gAPIClientId = strValue
      End If

      '
      ' Apply the API Secret Key
      '
      If strSection = "API" And strKey = "ClientSecret" Then
        If strValue.Length = 0 Then Exit Sub
        gAPIClientSecret = strValue
      End If

      '
      ' Apply the API Username
      '
      If strSection = "API" And strKey = "Username" Then
        gAPIUsername = strValue
      End If

      '
      ' Apply the API Password
      '
      If strSection = "API" And strKey = "Password" Then
        If strValue.Length = 0 Then Exit Sub
        gAPIPassword = strValue
        strValue = hs.EncryptString(strValue, "&Cul8r#1")
      End If

      '
      ' Apply the Sighthound SnapshotsMaxWidth
      '
      If strSection = "Options" And strKey = "SnapshotsMaxWidth" Then
        gSnapshotMaxWidth = strValue
      End If

      If strSection = "Options" And strKey = "SnapshotEventMaxWidth" Then
        gSnapshotEventMaxWidth = strValue
      End If

      If strSection = "Options" And strKey = "VideoQuality" Then
        gEventVideoQuality = strValue
      End If

      If strSection = "EmailNotification" And strKey = "EmailEnabled" Then
        gEventEmailNotification = CBool(strValue)
      End If

      '
      ' Save the settings
      '
      hs.SaveINISetting(strSection, strKey, strValue, gINIFile)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SaveSetting()")
    End Try

  End Sub

#End Region

#Region "UltraNetatmoWelcome3 Actions/Triggers/Conditions"

#Region "Trigger Proerties"

  ''' <summary>
  ''' Defines the valid triggers for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetTriggers()
    Dim o As Object = Nothing
    If triggers.Count = 0 Then
      triggers.Add(o, "Welcome Trigger")           ' 1
    End If
  End Sub

  ''' <summary>
  ''' Lets HomeSeer know our plug-in has triggers
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasTriggers() As Boolean
    Get
      SetTriggers()
      Return IIf(triggers.Count > 0, True, False)
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerCount() As Integer
    SetTriggers()
    Return triggers.Count
  End Function

  ''' <summary>
  ''' Returns the subtrigger count
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer
    Get
      Dim trigger As trigger
      If ValidTrig(TriggerNumber) Then
        trigger = triggers(TriggerNumber - 1)
        If Not (trigger Is Nothing) Then
          Return 0
        Else
          Return 0
        End If
      Else
        Return 0
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String
    Get
      If Not ValidTrig(TriggerNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, triggers.Keys(TriggerNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the subtrigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <param name="SubTriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String
    Get
      'Dim trigger As trigger
      If ValidSubTrig(TriggerNumber, SubTriggerNumber) Then
        Return ""
      Else
        Return ""
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is valid
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidTrig(ByVal TrigIn As Integer) As Boolean
    SetTriggers()
    If TrigIn > 0 AndAlso TrigIn <= triggers.Count Then
      Return True
    End If
    Return False
  End Function

  ''' <summary>
  ''' Determines if the trigger is a valid subtrigger
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <param name="SubTrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidSubTrig(ByVal TrigIn As Integer, ByVal SubTrigIn As Integer) As Boolean
    Return False
  End Function

  ''' <summary>
  ''' Tell HomeSeer which triggers have conditions
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean
    Get
      Select Case TriggerNumber
        Case 0
          Return True   ' Render trigger as IF / AND IF
        Case Else
          Return False  ' Render trigger as IF / OR IF
      End Select
    End Get
  End Property

  ''' <summary>
  ''' HomeSeer will set this to TRUE if the trigger is being used as a CONDITION.  
  ''' Check this value in BuildUI and other procedures to change how the trigger is rendered if it is being used as a condition or a trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Set(ByVal value As Boolean)

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      ' TriggerCondition(sKey) = value

    End Set
    Get

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Return False

    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is a condition
  ''' </summary>
  ''' <param name="sKey"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property TriggerCondition(sKey As String) As Boolean
    Get

      If conditions.ContainsKey(sKey) = True Then
        Return conditions(sKey)
      Else
        Return False
      End If

    End Get
    Set(value As Boolean)

      If conditions.ContainsKey(sKey) = False Then
        conditions.Add(sKey, value)
      Else
        conditions(sKey) = value
      End If

    End Set
  End Property

  ''' <summary>
  ''' Called when HomeSeer wants to check if a condition is true
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Return False
  End Function

#End Region

#Region "Trigger Interface"

  ''' <summary>
  ''' Builds the Trigger UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = TrigInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    Else 'new event, so clean out the trigger object
      trigger = New trigger
    End If

    stb.AppendLine("<table cellspacing='0'>")

    Select Case TrigInfo.TANumber
      Case NetatmoTriggers.WelcomeTriggers
        Dim triggerName As String = GetEnumName(NetatmoTriggers.WelcomeTriggers)

        '
        ' Start Welcome Trigger
        '
        Dim ActionSelected As String = trigger.Item("WelcomeTrigger")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", triggerName, "WelcomeTrigger", UID, sUnique)

        Dim jqWelcomeTrigger As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqWelcomeTrigger.autoPostBack = True

        jqWelcomeTrigger.AddItem("(Select Welcome Trigger)", "", (ActionSelected = ""))

        Dim itemValues As Array = System.Enum.GetValues(GetType(WelcomeTriggers))
        Dim itemNames As Array = System.Enum.GetNames(GetType(WelcomeTriggers))

        For i As Integer = 0 To itemNames.Length - 1
          Dim strOptionName = GetEnumDescription(CType(i, WelcomeTriggers))
          Dim strOptionValue = i
          jqWelcomeTrigger.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue.ToString))
        Next

        stb.AppendLine(" <tr>")
        stb.AppendLine("  <td class=""event_Txt_Selection"">Trigger:</td>")
        stb.AppendFormat("<td class=""event_Txt_Option"">{0}</td>", jqWelcomeTrigger.Build)
        stb.AppendLine(" </tr>")

        '
        ' Start Welcome Camera
        '
        ActionSelected = trigger.Item("WelcomeCamera")

        actionId = String.Format("{0}{1}_{2}_{3}", triggerName, "WelcomeCamera", UID, sUnique)

        Dim jqWelcomeCamera As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqWelcomeCamera.autoPostBack = True

        jqWelcomeCamera.AddItem("(Select Welcome Camera)", "", (ActionSelected = ""))
        jqWelcomeCamera.AddItem("Any Camera", "*", (ActionSelected = "*"))

        Dim WelcomeCameras As List(Of hspi_netatmo_api.Camera) = NetatmoAPI.GetWelcomeCameras
        For Each WelcomeCamera As hspi_netatmo_api.Camera In WelcomeCameras
          Dim camera_id As String = WelcomeCamera.id
          Dim camera_name As String = WelcomeCamera.name

          Dim strOptionValue As String = camera_id
          Dim strOptionName As String = camera_name
          jqWelcomeCamera.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.AppendLine(" <tr>")
        stb.AppendLine("  <td class=""event_Txt_Selection"">Welcome Camera:</td>")
        stb.AppendFormat("<td class=""event_Txt_Option"">{0}</td>", jqWelcomeCamera.Build)
        stb.AppendLine(" </tr>")

    End Select

    stb.AppendLine("</table>")

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Process changes to the trigger from the HomeSeer events page
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection,
                                       ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As String = TrigInfo.UID.ToString
    Dim TANumber As Integer = TrigInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = TrigInfo.DataIn
    Ret.TrigActInfo = TrigInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    ' DeSerializeObject
    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If
    trigger.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case NetatmoTriggers.WelcomeTriggers
          Dim triggerName As String = GetEnumName(NetatmoTriggers.WelcomeTriggers)

          For Each sKey As String In parts.Keys
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, triggerName & "WelcomeTrigger_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("WelcomeTrigger") = ActionValue

              Case InStr(sKey, triggerName & "WelcomeCamera_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("WelcomeCamera") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(trigger, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Trigger not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Trigger UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret

  End Function

  ''' <summary>
  ''' Determines if a trigger is properly configured
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Get
      Dim Configured As Boolean = True
      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Select Case TrigInfo.TANumber
        Case NetatmoTriggers.WelcomeTriggers
          If trigger.Item("WelcomeTrigger") = "" Then Configured = False
          If trigger.Item("WelcomeCamera") = "" Then Configured = False

      End Select

      Return Configured
    End Get
  End Property

  ''' <summary>
  ''' Formats the trigger for display
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim stb As New StringBuilder

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Select Case TrigInfo.TANumber
      Case NetatmoTriggers.WelcomeTriggers
        If trigger.uid <= 0 Then
          stb.Append("Trigger has not been properly configured.")
        Else
          Dim strTriggerName As String = GetEnumDescription(NetatmoTriggers.WelcomeTriggers)
          Dim strWelcomeTrigger As String = trigger.Item("WelcomeTrigger")
          Dim strWelcomeCamera As String = trigger.Item("WelcomeCamera")

          SyncLock gNetatmoCamerasLock
            If gNetatmoCameras.Any(Function(s) s.id = strWelcomeCamera) = True Then
              Dim NetatmoCamera As hspi_netatmo_api.Camera = gNetatmoCameras.Find(Function(s) s.id = strWelcomeCamera)
              strWelcomeCamera = NetatmoCamera.name
            End If
          End SyncLock

          If strWelcomeCamera = "*" Then strWelcomeCamera = "Any Camera"

          stb.AppendLine("<table cellspacing='0'>")
          stb.AppendLine(" <tr>")
          stb.AppendLine("  <td class=""event_if_then_text"" colspan=""2"">")
          stb.AppendFormat("{0}: {1}", IFACE_NAME, strTriggerName)
          stb.AppendLine("  </td>")
          stb.AppendLine(" </tr>")
          stb.AppendLine(" <tr>")
          stb.AppendLine("  <td class=""event_Txt_Selection"">Trigger:</td>")
          stb.AppendFormat("<td class=""event_Txt_Option"">{0}</td>", GetEnumDescription(CType(strWelcomeTrigger, WelcomeTriggers)))
          stb.AppendLine(" </tr>")
          stb.AppendLine(" <tr>")
          stb.AppendLine("  <td class=""event_Txt_Selection"">Welcome Camera:</td>")
          stb.AppendFormat("<td class=""event_Txt_Option"">{0}</td>", strWelcomeCamera)
          stb.AppendLine(" </tr>")
          stb.AppendLine("</table>")

        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Checks to see if trigger should fire
  ''' </summary>
  ''' <param name="Plug_Name"></param>
  ''' <param name="TrigID"></param>
  ''' <param name="SubTrig"></param>
  ''' <param name="strTrigger"></param>
  ''' <remarks></remarks>
  Private Sub CheckTrigger(Plug_Name As String, TrigID As Integer, SubTrig As Integer, strTrigger As String)

    Try
      '
      ' Check HomeSeer Triggers
      '
      If Plug_Name.Contains(":") = False Then Plug_Name &= ":"
      Dim TrigsToCheck() As IAllRemoteAPI.strTrigActInfo = callback.TriggerMatches(Plug_Name, TrigID, -1)

      Try

        For Each TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo In TrigsToCheck
          Dim UID As String = TrigInfo.UID.ToString

          If Not (TrigInfo.DataIn Is Nothing) Then

            Dim trigger As New trigger
            DeSerializeObject(TrigInfo.DataIn, trigger)

            Select Case TrigID

              Case NetatmoTriggers.WelcomeTriggers

                Dim strWelcomeTrigger As String = GetEnumName(CType(SubTrig, WelcomeTriggers))
                Dim strWelcomeCamera As String = trigger.Item("WelcomeCamera")

                Dim strTriggerCheck As String = String.Format("{0},{1}", strWelcomeTrigger, strWelcomeCamera)
                If Regex.IsMatch(strTrigger, strTriggerCheck) = True Then
                  callback.TriggerFire(IFACE_NAME, TrigInfo)
                End If

            End Select

          End If

        Next

      Catch pEx As Exception

      End Try

    Catch pEx As Exception

    End Try

  End Sub

#End Region

#Region "Action Properties"

  ''' <summary>
  ''' Defines the valid actions for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetActions()
    Dim o As Object = Nothing
    If actions.Count = 0 Then
      'actions.Add(o, "Email Notification")          ' 1
      'actions.Add(o, "Speak Weather")               ' 2
    End If
  End Sub

  ''' <summary>
  ''' Returns the action count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ActionCount() As Integer
    SetActions()
    Return actions.Count
  End Function

  ''' <summary>
  ''' Returns the action name
  ''' </summary>
  ''' <param name="ActionNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String
    Get
      If Not ValidAction(ActionNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, actions.Keys(ActionNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if an action is valid
  ''' </summary>
  ''' <param name="ActionIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidAction(ByVal ActionIn As Integer) As Boolean
    SetActions()
    If ActionIn > 0 AndAlso ActionIn <= actions.Count Then
      Return True
    End If
    Return False
  End Function

#End Region

#Region "Action Interface"

  ''' <summary>
  ''' Builds the Action UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks>This function is called from the HomeSeer event page when an event is in edit mode.</remarks>
  Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = ActInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case NetatmoActions.EmailNotification
        Dim actionName As String = GetEnumName(NetatmoActions.EmailNotification)

        '
        ' Start EmailNotification
        '
        Dim ActionSelected As String = action.Item("Notification")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "Notification", UID, sUnique)

        Dim jqNotifications As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqNotifications.autoPostBack = True

        jqNotifications.AddItem("(Select E-mail Notification)", "", (ActionSelected = ""))
        Dim Actions As String() = {"Weather Conditions", "Weather Forecast", "Weather Alerts"}
        For Each strAction As String In Actions
          Dim strOptionValue As String = strAction
          Dim strOptionName As String = strOptionValue
          jqNotifications.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append(jqNotifications.Build)

        '
        ' Start Station Name
        '
        ActionSelected = action.Item("Station")

        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "Station", UID, sUnique)

        Dim jqStation As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqStation.autoPostBack = True

        jqStation.AddItem("(Select Weather Station)", "", (ActionSelected = ""))

        For index As Integer = 1 To 6
          Dim strStationNumber As String = String.Format("Station{0}", index.ToString)
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", "Not defined", gINIFile))

          Dim strOptionValue As String = strStationNumber
          Dim strOptionName As String = String.Format("{0} [{1}]", strStationNumber, strStationName)
          jqStation.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append("from")
        stb.Append(jqStation.Build)

      Case NetatmoActions.SpeakWeather
        Dim actionName As String = GetEnumName(NetatmoActions.SpeakWeather)

        '
        ' Start Speak Weather
        '
        Dim ActionSelected As String = action.Item("Notification")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "Notification", UID, sUnique)

        Dim jqNotifications As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqNotifications.autoPostBack = True

        jqNotifications.AddItem("(Select Speak Action)", "", (ActionSelected = ""))
        Dim Actions As String() = {"Weather Conditions", "Weather Forecast", "Weather Alerts"}
        For Each strAction As String In Actions
          Dim strOptionValue As String = strAction
          Dim strOptionName As String = strOptionValue
          jqNotifications.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append(jqNotifications.Build)

        '
        ' Start Station Name
        '
        ActionSelected = action.Item("Station")

        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "Station", UID, sUnique)

        Dim jqStation As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqStation.autoPostBack = True

        jqStation.AddItem("(Select Weather Station)", "", (ActionSelected = ""))

        For index As Integer = 1 To 6
          Dim strStationNumber As String = String.Format("Station{0}", index.ToString)
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", "Not defined", gINIFile))

          Dim strOptionValue As String = strStationNumber
          Dim strOptionName As String = String.Format("{0} [{1}]", strStationNumber, strStationName)
          jqStation.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append("from")
        stb.Append(jqStation.Build)

        '
        ' Start Speaker Host
        '
        ActionSelected = IIf(action.Item("SpeakerHost").Length = 0, "*:*", action.Item("SpeakerHost"))
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "SpeakerHost", UID, sUnique)

        Dim jqSpeakerHost As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 45, True)
        stb.Append("Host(host:instance)")
        stb.Append(jqSpeakerHost.Build)

    End Select

    Return stb.ToString

  End Function

  ''' <summary>
  ''' When a user edits your event actions in the HomeSeer events, this function is called to process the selections.
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection,
                                      ByVal ActInfo As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As Integer = ActInfo.UID
    Dim TANumber As Integer = ActInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = ActInfo.DataIn
    Ret.TrigActInfo = ActInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    '
    ' DeSerializeObject
    '
    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If
    action.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case NetatmoActions.EmailNotification
          Dim actionName As String = GetEnumName(NetatmoActions.EmailNotification)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "Notification_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Notification") = ActionValue

              Case InStr(sKey, actionName & "Station_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Station") = ActionValue

            End Select
          Next

        Case NetatmoActions.SpeakWeather
          Dim actionName As String = GetEnumName(NetatmoActions.SpeakWeather)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "Notification_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Notification") = ActionValue

              Case InStr(sKey, actionName & "Station_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Station") = ActionValue

              Case InStr(sKey, actionName & "SpeakerHost_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("SpeakerHost") = ActionValue
            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(action, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Action not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Action UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret
  End Function

  ''' <summary>
  ''' Determines if our action is proplery configured
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return TRUE if the given action is configured properly</returns>
  ''' <remarks>There may be times when a user can select invalid selections for the action and in this case you would return FALSE so HomeSeer will not allow the action to be saved.</remarks>
  Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim Configured As Boolean = True
    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case NetatmoActions.EmailNotification
        If action.Item("Notification") = "" Then Configured = False
        If action.Item("Station") = "" Then Configured = False

      Case NetatmoActions.SpeakWeather
        If action.Item("Notification") = "" Then Configured = False
        If action.Item("Station") = "" Then Configured = False
        If action.Item("SpeakerHost") = "" Then Configured = False

    End Select

    Return Configured

  End Function

  ''' <summary>
  ''' After the action has been configured, this function is called in your plugin to display the configured action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return text that describes the given action.</returns>
  ''' <remarks></remarks>
  Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String
    Dim stb As New StringBuilder

    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber

      Case NetatmoActions.EmailNotification
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim strActionName = GetEnumDescription(NetatmoActions.EmailNotification)

          Dim strNotificationType As String = action.Item("Notification")

          Dim strStationNumber As String = action.Item("Station")
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", strStationNumber, gINIFile))

          If strStationName.Length > 0 Then
            strStationNumber = String.Format("{0} [{1}]", strStationNumber, strStationName)
          End If

          stb.AppendFormat("{0} {1} {2}", strActionName, strStationNumber, strNotificationType)
        End If

      Case NetatmoActions.SpeakWeather
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim strActionName = GetEnumDescription(NetatmoActions.SpeakWeather)

          Dim strNotificationType As String = action.Item("Notification")

          Dim strStationNumber As String = action.Item("Station")
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", strStationNumber, gINIFile))

          Dim strSpeakerHost As String = action.Item("SpeakerHost")

          If strStationName.Length > 0 Then
            strStationNumber = String.Format("{0} [{1}]", strStationNumber, strStationName)
          End If

          stb.AppendFormat("{0} {1} from {2} on {3}", strActionName, strNotificationType, strStationNumber, strSpeakerHost)
        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Handles the HomeSeer Event Action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = ActInfo.UID.ToString

    Try

      Dim action As New action
      If Not (ActInfo.DataIn Is Nothing) Then
        DeSerializeObject(ActInfo.DataIn, action)
      Else
        Return False
      End If

      Select Case ActInfo.TANumber

        Case NetatmoActions.EmailNotification
          Dim strNotificationType As String = action.Item("Notification")
          Dim strStationNumber As String = action.Item("Station")

          Select Case strNotificationType
            Case "Weather Conditions"
              'EmailWeatherConditions(strStationNumber)

            Case "Weather Forecast"
              'EmailWeatherForecast(strStationNumber)

            Case "Weather Alerts"
              'EmailWeatherAlerts(strStationNumber)

          End Select

        Case NetatmoActions.SpeakWeather
          Dim strNotificationType As String = action.Item("Notification")
          Dim strStationNumber As String = action.Item("Station")
          Dim strSpeakerHost As String = action.Item("SpeakerHost")

          Select Case strNotificationType
            Case "Weather Conditions"
              'SpeakWeatherConditions(strStationNumber, False, strSpeakerHost)

            Case "Weather Forecast"
              'SpeakWeatherForecast(strStationNumber, False, strSpeakerHost)

            Case "Weather Alerts"
              'SpeakWeatherAlerts(strStationNumber, False, strSpeakerHost)
          End Select

      End Select

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      hs.WriteLog(IFACE_NAME, "Error executing action: " & pEx.Message)
    End Try

    Return True

  End Function

#End Region

#End Region

End Module

Public Enum NetatmoTriggers
  <Description("Welcome Triggers")>
  WelcomeTriggers = 1
End Enum

Public Enum NetatmoActions
  <Description("Email Notification")>
  EmailNotification = 1
  <Description("Speak Weather")>
  SpeakWeather = 2
End Enum

<Flags()> Public Enum CameraStatus
  NotLocal = -3
  PowerSupply = -2
  SDCard = -1
  Offline = 0
  Online = 1
End Enum

<Flags()> Public Enum WelcomeTriggers
  <Description("Motion Has Been Detected")>
  MotionDetected = 0
  <Description("Someone Unknown Has Been Seen (Welcome)")>
  SomeoneUnknownSceen = 1
  <Description("Someone Known Has Been Seen (Welcome)")>
  SomeoneknownSceen = 2
  <Description("Someone Known Arrives Home (Welcome)")>
  SomeoneKnownArrives = 3
  <Description("Someone Known Leaves Home (Welcome)")>
  SomeoneKnownLeaves = 4
  <Description("Camera Monitoring Has Been Enabled (Welcome/Presence)")>
  MonitoringEnabled = 5
  <Description("Welcome Camera Monitoring Has Been Disabled (Welcome/Presence)")>
  MonitoringDisabled = 6
  <Description("Camera Connection Lost (Welcome/Presence)")>
  CameraDisconnected = 7
  <Description("Camera Connection Established (Welcome/Presence)")>
  CameraConnected = 8
  <Description("Camera Has Been Powered On (Welcome/Presence)")>
  WelcomeCameraBoot = 9
  <Description("Outdoor Event - Human Seen (Presence)")>
  OutdoorEventHumanSeen = 10
  <Description("Outdoor Event - Vehicle Seen (Presence)")>
  OutdoorEventVehicleSeen = 11
  <Description("Outdoor Event - Animal Seen (Presence)")>
  OutdoorEventAnimalSeen = 12
End Enum
