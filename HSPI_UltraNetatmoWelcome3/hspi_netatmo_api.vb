Imports System.Net
Imports System.Web.Script.Serialization
Imports System.IO
Imports System.Text
Imports System.Drawing

Public Class hspi_netatmo_api

  Private _access_token As String = String.Empty
  Private _refresh_token As String = String.Empty
  Private _scope As String = String.Empty
  Private _expires_in As Integer = 0
  Private _expire_in As Integer = 0
  Private _refreshed As New Stopwatch

  Private _querySuccess As ULong = 0
  Private _queryFailure As ULong = 0

  Public Sub New()

  End Sub

  Public Function QuerySuccessCount() As ULong
    Return _querySuccess
  End Function

  Public Function QueryFailureCount() As ULong
    Return _queryFailure
  End Function

  ''' <summary>
  ''' Determines if the API key is valid
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidAPIToken() As Boolean

    Try

      If gAPIClientId.Length = 0 Then
        Return False
      ElseIf gAPIClientSecret.Length = 0 Then
        Return False
      ElseIf NetatmoAPI.GetAccessToken.Length = 0 Then
        Return False
      Else
        Return True
      End If

    Catch pEx As Exception
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Get Access Token
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetAccessToken() As String

    Dim expires_in As Long = _refreshed.ElapsedMilliseconds / 1000

    If CheckCredentials() = False Then
      WriteMessage("Invalid API authentication information.  Please check plug-in options.", MessageType.Error)
    ElseIf _access_token.Length = 0 Then
      GetToken()
    ElseIf expires_in > _expires_in Then
      _access_token = String.Empty
      RefreshAccessToken()
    End If

    Return _access_token

  End Function

  ''' <summary>
  ''' Gets Accesss Token
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub ApproveAccess()

    Try

      Dim data As Byte() = New ASCIIEncoding().GetBytes(String.Format("grant_type={0}&client_id={1}&client_secret={2}&username={3}&password={4}&scope={5}", "password", gAPIClientId, gAPIClientSecret, gAPIUsername, gAPIPassword, gAPIScope))

      Dim strURL As String = String.Format("https://api.netatmo.com/oauth2/token")
      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)

      HTTPWebRequest.Timeout = 1000 * 60
      HTTPWebRequest.Method = "POST"
      HTTPWebRequest.ContentType = "application/x-www-form-urlencoded;charset=UTF-8"
      HTTPWebRequest.ContentLength = data.Length

      Dim myStream As Stream = HTTPWebRequest.GetRequestStream
      myStream.Write(data, 0, data.Length)

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())

          Dim JSONString As String = reader.ReadToEnd()
          Dim js As New JavaScriptSerializer()
          Dim OAuth20 As oauth_token = js.Deserialize(Of oauth_token)(JSONString)

          ' {"access_token":"53d5a8531977598b59a1cd7c|c5b15ee74615b2cfa4d0598611a8f984","refresh_token":"53d5a8531977598b59a1cd7c|7216e7448890d3523ed2dad3ff5a03e5","scope":["read_station"],"expires_in":10800,"expire_in":10800}

          _access_token = OAuth20.access_token
          _refresh_token = OAuth20.refresh_token
          _expires_in = OAuth20.expires_in
          _expire_in = OAuth20.expire_in

          _refreshed.Start()

        End Using

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Debug)

      _access_token = String.Empty
      _refresh_token = String.Empty
      _expires_in = 0
      _expire_in = 0

    End Try

  End Sub

  ''' <summary>
  ''' Gets Accesss Token
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub GetToken()

    Try

      Dim data As Byte() = New ASCIIEncoding().GetBytes(String.Format("grant_type={0}&client_id={1}&client_secret={2}&username={3}&password={4}&scope={5}", "password", gAPIClientId, gAPIClientSecret, gAPIUsername, gAPIPassword, gAPIScope))

      Dim strURL As String = String.Format("https://api.netatmo.com/oauth2/token")
      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)

      HTTPWebRequest.Timeout = 1000 * 60
      HTTPWebRequest.Method = "POST"
      HTTPWebRequest.ContentType = "application/x-www-form-urlencoded;charset=UTF-8"
      HTTPWebRequest.ContentLength = data.Length

      Dim myStream As Stream = HTTPWebRequest.GetRequestStream
      myStream.Write(data, 0, data.Length)

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())

          Dim JSONString As String = reader.ReadToEnd()
          Dim js As New JavaScriptSerializer()
          Dim OAuth20 As oauth_token = js.Deserialize(Of oauth_token)(JSONString)

          ' {"access_token":"53d5a8531977598b59a1cd7c|c5b15ee74615b2cfa4d0598611a8f984","refresh_token":"53d5a8531977598b59a1cd7c|7216e7448890d3523ed2dad3ff5a03e5","scope":["read_station"],"expires_in":10800,"expire_in":10800}

          _access_token = OAuth20.access_token
          _refresh_token = OAuth20.refresh_token
          _expires_in = OAuth20.expires_in
          _expire_in = OAuth20.expire_in

          _refreshed.Start()

        End Using

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Debug)

      _access_token = String.Empty
      _refresh_token = String.Empty
      _expires_in = 0
      _expire_in = 0

    End Try

  End Sub

  ''' <summary>
  ''' Checks to see if required credentials are available
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function CheckCredentials() As Boolean

    Try

      Dim sbWarning As New StringBuilder

      If gAPIClientId.Length = 0 Then
        sbWarning.Append("Netatmo Client Id")
      End If
      If gAPIClientSecret.Length = 0 Then
        sbWarning.Append("Netatmo Client Secret")
      End If
      If gAPIUsername.Length = 0 Then
        sbWarning.Append("Netatmo Client Username")
      End If
      If gAPIPassword.Length = 0 Then
        sbWarning.Append("Netatmo Client Password")
      End If
      If sbWarning.Length = 0 Then Return True

    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Debug)
    End Try

    Return False

  End Function

  ''' <summary>
  ''' Refreshes the Access Token
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub RefreshAccessToken()

    Try

      Dim data As Byte() = New ASCIIEncoding().GetBytes(String.Format("grant_type={0}&client_id={1}&client_secret={2}&refresh_token={3}", "refresh_token", gAPIClientId, gAPIClientSecret, _refresh_token))

      Dim strURL As String = String.Format("https://api.netatmo.com/oauth2/token")
      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)

      HTTPWebRequest.Timeout = 1000 * 60
      HTTPWebRequest.Method = "POST"
      HTTPWebRequest.ContentType = "application/x-www-form-urlencoded;charset=UTF-8"
      HTTPWebRequest.ContentLength = data.Length

      Dim myStream As Stream = HTTPWebRequest.GetRequestStream
      myStream.Write(data, 0, data.Length)

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())

          Dim JSONString As String = reader.ReadToEnd()
          Dim js As New JavaScriptSerializer()
          Dim OAuth20 As refresh_token = js.Deserialize(Of refresh_token)(JSONString)

          ' {"access_token":"53d5a8531977598b59a1cd7c|c5b15ee74615b2cfa4d0598611a8f984","refresh_token":"53d5a8531977598b59a1cd7c|7216e7448890d3523ed2dad3ff5a03e5","scope":["read_station"],"expires_in":10800,"expire_in":10800}

          _access_token = OAuth20.access_token
          _refresh_token = OAuth20.refresh_token
          _expires_in = OAuth20.expires_in

          _refreshed.Start()

        End Using

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Debug)

      _access_token = String.Empty
      _refresh_token = String.Empty
      _expires_in = 0
      _expire_in = 0

    End Try

  End Sub

  ''' <summary>
  ''' Returns the Welcome Camera List
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWelcomeCameras() As List(Of hspi_netatmo_api.Camera)

    Try

      Return gNetatmoCameras

    Catch pEx As Exception
      Return Nothing
    End Try

  End Function

  ''' <summary>
  ''' Returns the Welcome Camera List
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWelcomeEvents() As List(Of hspi_netatmo_api.Event)

    Try

      Return gNetatmoEvents

    Catch pEx As Exception
      Return Nothing
    End Try

  End Function

  ''' <summary>
  ''' Gets the Welcome Home Data
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetHomeData() As HomeData

    Dim HomeData As New HomeData

    Try

      Dim access_token As String = Me.GetAccessToken()
      If access_token.Length = 0 Then
        Throw New Exception("Unable to get Netatmo Access Token.")
      End If
      Dim home_id As String = String.Empty

      Dim strURL As String = String.Format("https://api.netatmo.com/api/gethomedata?access_token={0}&size={1}", access_token, gEventSnapshotMax)
      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          HomeData = js.Deserialize(Of HomeData)(JSONString)

        End Using

      End Using

      _querySuccess += 1
    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Warning)

      _queryFailure += 1

      _expires_in = 0
    End Try

    Return HomeData

  End Function

  ''' <summary>
  ''' Gets the Welcome Local URL
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SendWelcomePing(ByVal vpn_url As String) As WelcomePing

    Dim WelcomePing As New WelcomePing

    Try

      If vpn_url.Length = 0 Then
        Throw New Exception("Unable to send ping to Welcome Camera because the vpn_url was empty.")
      End If
      Dim local_url As String = String.Empty

      Dim strURL As String = String.Format("{0}/command/ping", vpn_url)
      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          WelcomePing = js.Deserialize(Of WelcomePing)(JSONString)

        End Using

      End Using

      _querySuccess += 1
    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Warning)

      _queryFailure += 1

      _expires_in = 0
    End Try

    Return WelcomePing

  End Function

  ''' <summary>
  ''' Sets Persons Away
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SetPersonsAway(ByVal homes_id As String, ByVal person_id As String) As NetatmoResponse

    Dim NetatmoResponse As New NetatmoResponse

    Try

      Dim access_token As String = Me.GetAccessToken()
      If access_token.Length = 0 Then
        _expires_in = 0
        Throw New Exception("Unable to get Netatmo Access Token.")
      End If

      Dim strURL As String = String.Format("https://api.netatmo.com/api/setpersonsaway?access_token={0}&home_id={1}&person_id={2}", access_token, homes_id, person_id)
      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          NetatmoResponse = js.Deserialize(Of NetatmoResponse)(JSONString)

        End Using

      End Using

      _querySuccess += 1
    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Warning)

      _queryFailure += 1
    End Try

    Return NetatmoResponse

  End Function

  ''' <summary>
  ''' Sets Persons Home
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SetPersonsHome(ByVal homes_id As String, ByVal person_id As String) As NetatmoResponse

    Dim NetatmoResponse As New NetatmoResponse

    Try

      Dim access_token As String = Me.GetAccessToken()
      If access_token.Length = 0 Then
        _expires_in = 0
        Throw New Exception("Unable to get Netatmo Access Token.")
      End If

      Dim strURL As String = String.Format("https://api.netatmo.com/api/setpersonshome?access_token={0}&home_id={1}&person_ids[]={2}", access_token, homes_id, person_id)
      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          NetatmoResponse = js.Deserialize(Of NetatmoResponse)(JSONString)

        End Using

      End Using

      _querySuccess += 1
    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Warning)

      _queryFailure += 1
    End Try

    Return NetatmoResponse

  End Function

  ''' <summary>
  ''' You can also retrieve a real-time snapshot. To do so, add "/live/snapshot_720.jpg" to your camera URL.
  ''' </summary>
  ''' <param name="camera_url"></param>
  ''' <param name="strSnapshotFilename"></param>
  ''' <param name="strThumbnailFilename"></param>
  ''' <param name="timeout"></param>
  ''' <returns></returns>
  Public Function GetLiveSnapshot(ByVal camera_url As String,
                                  ByVal strSnapshotFilename As String,
                                  ByVal strThumbnailFilename As String,
                                  Optional timeout As Integer = 30) As WebExceptionStatus

    Try
      '
      ' Format the URL
      '
      Dim resource_url As String = String.Format("{0}/live/snapshot_720.jpg", camera_url)
      WriteMessage(String.Format("GetLiveSnapshot is running command: {0}.", resource_url), MessageType.Debug)

      '
      ' Build the HTTP Web Request
      '
      Dim lxRequest As HttpWebRequest = DirectCast(WebRequest.Create(resource_url), HttpWebRequest)
      lxRequest.Timeout = timeout * 1000

      '
      ' Process the HTTP Web Response
      '
      Using lxResponse As HttpWebResponse = DirectCast(lxRequest.GetResponse(), HttpWebResponse)

        Dim lnBuffer As Byte()
        Dim lnFile As Byte()

        Using lxBR As New BinaryReader(lxResponse.GetResponseStream())

          Using lxMS As New MemoryStream()

            lnBuffer = lxBR.ReadBytes(1024)

            While lnBuffer.Length > 0
              lxMS.Write(lnBuffer, 0, lnBuffer.Length)
              lnBuffer = lxBR.ReadBytes(1024)
            End While

            lnFile = New Byte(CInt(lxMS.Length) - 1) {}
            lxMS.Position = 0
            lxMS.Read(lnFile, 0, lnFile.Length)

            Try

              Using image As Image = Image.FromStream(lxMS)

                If Not image Is Nothing Then

                  If strSnapshotFilename.EndsWith("_snapshot.jpg") Then
                    If File.Exists(strSnapshotFilename) = True Then
                      File.SetCreationTime(strSnapshotFilename, DateTime.Now)
                    End If
                  End If

                  If strThumbnailFilename.EndsWith("_thumbnail.jpg") Then
                    If File.Exists(strThumbnailFilename) = True Then
                      File.SetCreationTime(strThumbnailFilename, DateTime.Now)
                    End If
                  End If

                  image.Save(strSnapshotFilename, image.RawFormat)

                  If strThumbnailFilename.Length > 0 Then
                    Dim iWidth As Integer = image.Width
                    Dim iHeight As Integer = image.Height
                    For i = 1 To 10 Step 1
                      iWidth = image.Width / i
                      iHeight = image.Height / i
                      If iHeight <= 36 Then Exit For
                    Next

                    Dim imageThumb As Image = image.GetThumbnailImage(iWidth, iHeight, Nothing, New IntPtr())
                    imageThumb.Save(strThumbnailFilename, image.RawFormat)
                  End If

                End If

                image.Dispose()

              End Using

            Catch pEx As ArgumentException
              '
              ' We got here because the data was not an image
              '
              Dim strErrorMessage As String = String.Format("The GetLiveSnapshot request sent to {0} failed: {1}", resource_url, pEx.Message)
              WriteMessage(strErrorMessage, MessageType.Warning)

            End Try

            lxMS.Close()
            lxBR.Close()

          End Using

        End Using

        lxResponse.Close()

      End Using

      Return WebExceptionStatus.Success

    Catch pEx As System.Net.WebException
      '
      ' Process the WebException
      '
      Dim strErrorMessage As String = String.Format("The GetLiveSnapshot request for image {0} failed: {1}", camera_url, pEx.Message)
      WriteMessage(strErrorMessage, MessageType.Debug)

      Return pEx.Status
    Catch pEx As Exception
      '
      ' Process the error
      '
      Dim strErrorMessage As String = String.Format("The GetLiveSnapshot request for image {0} failed: {1}", camera_url, pEx.Message)
      WriteMessage(strErrorMessage, MessageType.Debug)

      Return WebExceptionStatus.UnknownError
    End Try

  End Function

  ''' <summary>
  ''' You can also retrieve a real-time snapshot. To do so, add "/live/snapshot_720.jpg" to your camera URL.
  ''' </summary>
  ''' <param name="camera_url"></param>
  ''' <param name="strSnapshotFilename"></param>
  ''' <param name="strThumbnailFilename"></param>
  ''' <param name="timeout"></param>
  ''' <returns></returns>
  Public Function GetEventSnapshot(ByVal camera_url As String,
                                   ByVal picture_url As String,
                                   ByVal strSnapshotFilename As String,
                                   ByVal strThumbnailFilename As String,
                                   Optional timeout As Integer = 30) As WebExceptionStatus

    Try
      '
      ' Format the URL
      '
      Dim resource_url As String = String.Format("{0}/{1}", camera_url, picture_url.Replace("\", ""))
      WriteMessage(String.Format("GetEventSnapshot is running command: {0}.", resource_url), MessageType.Debug)

      '
      ' Build the HTTP Web Request
      '
      Dim lxRequest As HttpWebRequest = DirectCast(WebRequest.Create(resource_url), HttpWebRequest)
      lxRequest.Timeout = timeout * 1000

      '
      ' Process the HTTP Web Response
      '
      Using lxResponse As HttpWebResponse = DirectCast(lxRequest.GetResponse(), HttpWebResponse)

        Dim lnBuffer As Byte()
        Dim lnFile As Byte()

        Using lxBR As New BinaryReader(lxResponse.GetResponseStream())

          Using lxMS As New MemoryStream()

            lnBuffer = lxBR.ReadBytes(1024)

            While lnBuffer.Length > 0
              lxMS.Write(lnBuffer, 0, lnBuffer.Length)
              lnBuffer = lxBR.ReadBytes(1024)
            End While

            lnFile = New Byte(CInt(lxMS.Length) - 1) {}
            lxMS.Position = 0
            lxMS.Read(lnFile, 0, lnFile.Length)

            Try

              Using image As Image = Image.FromStream(lxMS)

                If Not image Is Nothing Then

                  If strSnapshotFilename.EndsWith("_snapshot.jpg") Then
                    If File.Exists(strSnapshotFilename) = True Then
                      File.SetCreationTime(strSnapshotFilename, DateTime.Now)
                    End If
                  End If

                  If strThumbnailFilename.EndsWith("_thumbnail.jpg") Then
                    If File.Exists(strThumbnailFilename) = True Then
                      File.SetCreationTime(strThumbnailFilename, DateTime.Now)
                    End If
                  End If

                  image.Save(strSnapshotFilename, image.RawFormat)

                  If strThumbnailFilename.Length > 0 Then
                    Dim iWidth As Integer = image.Width
                    Dim iHeight As Integer = image.Height
                    For i = 1 To 10 Step 1
                      iWidth = image.Width / i
                      iHeight = image.Height / i
                      If iHeight <= 36 Then Exit For
                    Next

                    Dim imageThumb As Image = image.GetThumbnailImage(iWidth, iHeight, Nothing, New IntPtr())
                    imageThumb.Save(strThumbnailFilename, image.RawFormat)
                  End If

                End If

                image.Dispose()

              End Using

            Catch pEx As ArgumentException
              '
              ' We got here because the data was not an image
              '
              Dim strErrorMessage As String = String.Format("The GetEventSnapshot request sent to {0} failed: {1}", resource_url, pEx.Message)
              WriteMessage(strErrorMessage, MessageType.Warning)

            End Try

            lxMS.Close()
            lxBR.Close()

          End Using

        End Using

        lxResponse.Close()

      End Using

      Return WebExceptionStatus.Success

    Catch pEx As System.Net.WebException
      '
      ' Process the WebException
      '
      Dim strErrorMessage As String = String.Format("The GetEventSnapshot request for image {0} failed: {1}", camera_url, pEx.Message)
      WriteMessage(strErrorMessage, MessageType.Debug)

      Return pEx.Status
    Catch pEx As Exception
      '
      ' Process the error
      '
      Dim strErrorMessage As String = String.Format("The GetEventSnapshot request for image {0} failed: {1}", camera_url, pEx.Message)
      WriteMessage(strErrorMessage, MessageType.Debug)

      Return WebExceptionStatus.UnknownError
    End Try

  End Function

  ''' <summary>
  ''' The method GetCameraPicture returns the associated picture : either a person's face or an event's snapshot
  ''' </summary>
  ''' <param name="image_id"></param>
  ''' <param name="key"></param>
  ''' <param name="timeout"></param>
  ''' <returns></returns>
  Public Function GetCameraPicture(ByVal image_id As String,
                                   ByVal key As String,
                                   ByVal strSnapshotFilename As String,
                                   ByVal strThumbnailFilename As String,
                                   Optional timeout As Integer = 30) As WebExceptionStatus

    Try
      '
      ' Format the URL
      '
      Dim resource_url As String = String.Format("https://api.netatmo.com/api/getcamerapicture?image_id={0}&key={1}", image_id, key)
      WriteMessage(String.Format("GetCameraPicture is running command: {0}.", resource_url), MessageType.Debug)

      '
      ' Build the HTTP Web Request
      '
      Dim lxRequest As HttpWebRequest = DirectCast(WebRequest.Create(resource_url), HttpWebRequest)
      lxRequest.Timeout = timeout * 1000

      '
      ' Process the HTTP Web Response
      '
      Using lxResponse As HttpWebResponse = DirectCast(lxRequest.GetResponse(), HttpWebResponse)

        Dim lnBuffer As Byte()
        Dim lnFile As Byte()

        Using lxBR As New BinaryReader(lxResponse.GetResponseStream())

          Using lxMS As New MemoryStream()

            lnBuffer = lxBR.ReadBytes(1024)

            While lnBuffer.Length > 0
              lxMS.Write(lnBuffer, 0, lnBuffer.Length)
              lnBuffer = lxBR.ReadBytes(1024)
            End While

            lnFile = New Byte(CInt(lxMS.Length) - 1) {}
            lxMS.Position = 0
            lxMS.Read(lnFile, 0, lnFile.Length)

            Try

              Using image As Image = Image.FromStream(lxMS)

                If Not image Is Nothing Then

                  If strSnapshotFilename.EndsWith("_snapshot.jpg") Then
                    If File.Exists(strSnapshotFilename) = True Then
                      File.SetCreationTime(strSnapshotFilename, DateTime.Now)
                    End If
                  End If

                  If strThumbnailFilename.EndsWith("_thumbnail.jpg") Then
                    If File.Exists(strThumbnailFilename) = True Then
                      File.SetCreationTime(strThumbnailFilename, DateTime.Now)
                    End If
                  End If

                  image.Save(strSnapshotFilename, image.RawFormat)

                  If strThumbnailFilename.Length > 0 Then
                    Dim iWidth As Integer = image.Width
                    Dim iHeight As Integer = image.Height
                    For i = 1 To 10 Step 1
                      iWidth = image.Width / i
                      iHeight = image.Height / i
                      If iHeight <= 36 Then Exit For
                    Next

                    Dim imageThumb As Image = image.GetThumbnailImage(iWidth, iHeight, Nothing, New IntPtr())
                    imageThumb.Save(strThumbnailFilename, image.RawFormat)
                  End If

                End If

                image.Dispose()

              End Using

            Catch pEx As ArgumentException
              '
              ' We got here because the data was not an image
              '
              Dim strErrorMessage As String = String.Format("The GetCameraPicture request sent to {0} failed: {1}", resource_url, pEx.Message)
              WriteMessage(strErrorMessage, MessageType.Warning)

            End Try

            lxMS.Close()
            lxBR.Close()

          End Using

        End Using

        lxResponse.Close()

      End Using

      Return WebExceptionStatus.Success

    Catch pEx As System.Net.WebException
      '
      ' Process the WebException
      '
      Dim strErrorMessage As String = String.Format("The GetCameraPicture request for image {0} failed: {1}", image_id, pEx.Message)
      WriteMessage(strErrorMessage, MessageType.Debug)

      Return pEx.Status
    Catch pEx As Exception
      '
      ' Process the error
      '
      Dim strErrorMessage As String = String.Format("The GetCameraPicture request for image {0} failed: {1}", image_id, pEx.Message)
      WriteMessage(strErrorMessage, MessageType.Debug)

      Return WebExceptionStatus.UnknownError
    End Try

  End Function

#Region "Netatmo oAuth Token"

  '{
  '    "access_token": "53d5a8531977598b59a1cd7c|c5b15ee74615b2cfa4d0598611a8f984",
  '    "refresh_token": "53d5a8531977598b59a1cd7c|7216e7448890d3523ed2dad3ff5a03e5",
  '    "scope": [
  '        "read_station",
  '        "read_thermostat",
  '        "write_thermostat"
  '    ],
  '    "expires_in": 10800,
  '    "expire_in": 10800
  '}

  <Serializable()>
  Private Class oauth_token
    Public Property access_token As String
    Public Property refresh_token As String
    Public Property scope As String()
    Public Property expires_in As Integer
    Public Property expire_in As Integer
  End Class

  <Serializable()>
  Private Class refresh_token
    Public Property access_token As String
    Public Property refresh_token As String
    Public Property expires_in As Integer
  End Class

#End Region

#Region "Welcome HomeData"

  <Serializable()>
  Public Class HomeData
    Public Property status As String = String.Empty
    Public Property body As HomesAndUsers
    Public Property time_exec As String = String.Empty
    Public Property time_server As Integer? = 0
  End Class

  <Serializable()>
  Public Class HomesAndUsers
    Public Property homes As List(Of Home)
    Public Property user As User
  End Class

  <Serializable()>
  Public Class Home
    Public Property id As String = String.Empty
    Public Property name As String = String.Empty
    Public Property place As Place
    Public Property persons As List(Of Person)
    Public Property events As List(Of [Event])
    Public Property cameras As List(Of Camera)
  End Class

  <Serializable()>
  Public Class Place
    Public Property city As String = String.Empty
    Public Property country As String = String.Empty
    Public Property timezone As String = String.Empty
  End Class

  <Serializable()>
  Public Class Person
    Public Property id As String = String.Empty
    Public Property last_seen As String = String.Empty
    Public Property out_of_sight As String = String.Empty
    Public Property face As face
    Public Property pseudo As String = String.Empty
  End Class

  <Serializable()>
  Public Class face
    Public Property id As String = String.Empty
    Public Property version As String = String.Empty
    Public Property key As String = String.Empty
  End Class

  <Serializable()>
  Public Class [Event]
    Public Property id As String = String.Empty
    Public Property type As String = String.Empty
    Public Property time As Integer? = 0
    Public Property offset As Integer? = 0
    Public Property camera_id As String = String.Empty
    Public Property person_id As String = String.Empty
    Public Property snapshot As Snapshot
    Public Property snapshot_downloaded As Boolean = False

    Public Property video_id As String = String.Empty
    Public Property video_status As String = String.Empty
    Public Property event_list As List(Of Outdoor_Event)
    Public Property is_arrival As String = String.Empty
    Public Property message As String = String.Empty
    Public Property sub_message As String = String.Empty

    Public Property video_downloaded As Boolean = False
    Public Property download_tries As Integer = 0
  End Class

  <Serializable()>
  Public Class Outdoor_Event
    Public Property id As String = String.Empty
    Public Property type As String = String.Empty
    Public Property time As Integer? = 0
    Public Property offset As Integer? = 0
    Public Property message As String = String.Empty
    Public Property snapshot As Snapshot
    Public Property snapshot_downloaded As Boolean = False
    Public Property video_id As String = String.Empty
    Public Property video_status As String = String.Empty

    Public Property video_downloaded As Boolean = False
    Public Property download_tries As Integer = 0
  End Class

  <Serializable()>
  Public Class Snapshot
    Public Property id As String = String.Empty
    Public Property version As String = String.Empty
    Public Property key As String = String.Empty
    Public Property filename As String = String.Empty
  End Class

  <Serializable()>
  Public Class Camera
    Public Property id As String = String.Empty
    Public Property status As String = String.Empty
    Public Property vpn_url As String = String.Empty
    Public Property is_local As Boolean? = False
    Public Property sd_status As String = String.Empty
    Public Property alim_status As String = String.Empty
    Public Property name As String = String.Empty
    Public Property camera_url As String = String.Empty
    Public Property light_mode_status As String = String.Empty
  End Class

  <Serializable()>
  Public Class User
    Public Property reg_locale As String = String.Empty
    Public Property lang As String = String.Empty
    Public Property country As String = String.Empty
  End Class

#End Region

#Region "Welcome Ping"
  <Serializable()>
  Public Class WelcomePing
    Public Property local_url As String = String.Empty
    Public Property product_name As String = String.Empty
  End Class
#End Region

#Region "Generic Netatmo Response"
  <Serializable()>
  Public Class NetatmoResponse
    Public Property status As String = String.Empty
    Public Property time_exec As Double = 0
    Public Property time_server As Integer = 0
  End Class
#End Region
End Class
