﻿Imports System.IO
Imports System.Runtime.Serialization.Formatters
Imports System.Text
Imports System.Security.Cryptography

Module hspi_utils

  Public Function GetTimeDiff(startTime As Long, endTime As DateTime) As String

    Dim sb As New StringBuilder

    Try

      Dim lastSceen As DateTime = ConvertEpochToDateTime(startTime)
      Dim timeSpan As TimeSpan = endTime - lastSceen

      Dim years As Integer = Val(timeSpan.Days / 365)
      Dim months As Integer = Val(timeSpan.Days / 30.436875)
      Dim days As Integer = timeSpan.Days
      Dim hours As Integer = timeSpan.Hours
      Dim minutes As Integer = timeSpan.Minutes

      'If years > 0 Then
      '  sb.AppendFormat("{0}{1} {2}", IIf(sb.Length > 0, ", ", ""), years.ToString, IIf(years = 1, "year", "years"))
      'End If
      'If months > 0 Then
      '  sb.AppendFormat("{0}{1} {2}", IIf(sb.Length > 0, ", ", ""), months.ToString, IIf(months = 1, "month", "months"))
      'End If
      If days > 0 Then
        sb.AppendFormat("{0}{1} {2}", IIf(sb.Length > 0, ", ", ""), days.ToString, IIf(days = 1, "day", "days"))
      End If

      If sb.Length = 0 Then
        If hours > 0 Then
          sb.AppendFormat("{0}{1} {2}", IIf(sb.Length > 0, ", ", ""), hours.ToString, IIf(hours = 1, "hour", "hours"))
        End If
        sb.AppendFormat("{0}{1} {2}", IIf(sb.Length > 0, ", ", ""), minutes.ToString, IIf(minutes = 1, "minute", "minutes"))
      End If

    Catch pEx As Exception
      Return ""
    End Try

    Return sb.ToString

  End Function

  ''' <summary>
  ''' Fixes the filename path
  ''' </summary>
  ''' <param name="strFileName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function FixPath(ByVal strFileName As String) As String

    Try

      Dim OSType As HomeSeerAPI.eOSType = hs.GetOSType()

      If OSType = HomeSeerAPI.eOSType.linux Then

        strFileName = strFileName.Replace("\", "/")
        strFileName = strFileName.Replace("//", "/")

      Else

        strFileName = strFileName.Replace("/", "\")
        strFileName = strFileName.Replace("\\", "\")

      End If

    Catch pEx As Exception

    End Try

    Return strFileName

  End Function

  ''' <summary>
  ''' Generates hash 
  ''' </summary>
  ''' <param name="SourceText"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GenerateHash(ByVal SourceText As String) As String

    Try
      '
      ' Create an encoding object to ensure the encoding standard for the source text
      '
      Dim Ue As New UnicodeEncoding()

      '
      ' Retrieve a byte array based on the source text
      '
      Dim ByteSourceText() As Byte = Ue.GetBytes(SourceText)

      '
      ' Instantiate an MD5 Provider object
      '
      Dim Md5 As New MD5CryptoServiceProvider()

      '
      ' Compute the hash value from the source
      '
      Dim ByteHash() As Byte = Md5.ComputeHash(ByteSourceText)

      '
      ' And convert it to String format for return
      '
      Return Convert.ToBase64String(ByteHash)

    Catch ex As Exception
      '
      ' Ignore this error
      '
      Return Len(SourceText).ToString
    End Try

  End Function

  ''' <summary>
  ''' Registers a web page with HomeSeer
  ''' </summary>
  ''' <param name="link"></param>
  ''' <param name="linktext"></param>
  ''' <param name="page_title"></param>
  ''' <param name="Instance"></param>
  ''' <remarks></remarks>
  Public Sub RegisterWebPage(ByVal link As String,
                             Optional linktext As String = "",
                             Optional page_title As String = "",
                             Optional Instance As String = "")

    Try

      hs.RegisterPage(link, IFACE_NAME, Instance)

      If linktext = "" Then linktext = link
      linktext = linktext.Replace("_", " ")

      If page_title = "" Then page_title = linktext

      Dim wpd As New HomeSeerAPI.WebPageDesc
      wpd.plugInName = IFACE_NAME
      wpd.plugInInstance = Instance
      wpd.link = link
      wpd.linktext = linktext & Instance
      wpd.page_title = page_title & Instance

      callback.RegisterLink(wpd)

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "RegisterWebPage")

    End Try

  End Sub

  ''' <summary>
  ''' Registers links to AXPX web page
  ''' </summary>
  ''' <param name="link"></param>
  ''' <param name="linktext"></param>
  ''' <param name="page_title"></param>
  ''' <param name="Instance"></param>
  ''' <remarks></remarks>
  Public Sub RegisterASXPWebPage(ByVal link As String,
                                 linktext As String,
                                 page_title As String,
                                 Optional Instance As String = "")

    Try

      Dim wpd As New HomeSeerAPI.WebPageDesc

      wpd.plugInName = IFACE_NAME
      wpd.plugInInstance = Instance
      wpd.link = link
      wpd.linktext = linktext & Instance
      wpd.page_title = page_title & Instance

      callback.RegisterLink(wpd)

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "RegisterASXPWebPage")

    End Try

  End Sub

  ''' <summary>
  ''' Registers the Help File URL
  ''' </summary>
  ''' <param name="link"></param>
  ''' <param name="linktext"></param>
  ''' <param name="page_title"></param>
  ''' <param name="Instance"></param>
  ''' <remarks></remarks>
  Public Sub RegisterHelpPage(ByVal link As String,
                              linktext As String,
                              page_title As String,
                              Optional Instance As String = "")

    Try

      Dim wpd As New HomeSeerAPI.WebPageDesc

      wpd.plugInName = IFACE_NAME
      wpd.plugInInstance = Instance
      wpd.link = link
      wpd.linktext = linktext & Instance
      wpd.page_title = page_title & Instance

      hs.RegisterHelpLink(wpd)

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "RegisterHelpPage")

    End Try

  End Sub

  ''' <summary>
  ''' Serialize Object
  ''' </summary>
  ''' <param name="ObjIn"></param>
  ''' <param name="bteOut"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SerializeObject(ByRef ObjIn As Object, ByRef bteOut() As Byte) As Boolean
    If ObjIn Is Nothing Then Return False
    Dim str As New MemoryStream
    Dim sf As New Binary.BinaryFormatter

    Try
      sf.Serialize(str, ObjIn)
      ReDim bteOut(CInt(str.Length - 1))
      bteOut = str.ToArray
      Return True
    Catch pEx As Exception
      WriteMessage(IFACE_NAME & " Error: Serializing object " & ObjIn.ToString & " :" & pEx.Message, MessageType.Critical)
      Return False
    End Try

  End Function

  ''' <summary>
  ''' DeSerialize Object
  ''' </summary>
  ''' <param name="bteIn"></param>
  ''' <param name="ObjOut"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DeSerializeObject(ByRef bteIn() As Byte, ByRef ObjOut As Object) As Boolean

    ' Almost immediately there is a test to see if ObjOut is NOTHING.  The reason for this
    '   when the ObjOut is suppose to be where the deserialized object is stored, is that 
    '   I could find no way to test to see if the deserialized object and the variable to 
    '   hold it was of the same type.  If you try to get the type of a null object, you get
    '   only a null reference exception!  If I do not test the object type beforehand and 
    '   there is a difference, then the InvalidCastException is thrown back in the CALLING
    '   procedure, not here, because the cast is made when the ByRef object is cast when this
    '   procedure returns, not earlier.  In order to prevent a cast exception in the calling
    '   procedure that may or may not be handled, I made it so that you have to at least 
    '   provide an initialized ObjOut when you call this - ObjOut is set to nothing after it 
    '   is typed.

    If bteIn Is Nothing Then Return False
    If bteIn.Length < 1 Then Return False
    If ObjOut Is Nothing Then Return False

    Dim str As MemoryStream
    Dim sf As New Binary.BinaryFormatter
    Dim ObjTest As Object
    Dim TType As System.Type
    Dim OType As System.Type

    Try
      OType = ObjOut.GetType
      ObjOut = Nothing
      str = New MemoryStream(bteIn)
      ObjTest = sf.Deserialize(str)
      If ObjTest Is Nothing Then Return False
      TType = ObjTest.GetType
      'If Not TType.Equals(OType) Then Return False
      ObjOut = ObjTest
      If ObjOut Is Nothing Then Return False
      Return True
    Catch exIC As InvalidCastException
      Return False
    Catch ex As Exception
      WriteMessage(IFACE_NAME & " Error: Serializing object " & ex.Message, MessageType.Critical)
      Return False
    End Try

  End Function

End Module
