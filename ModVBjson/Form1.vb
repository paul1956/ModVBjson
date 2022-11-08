Imports System.IO
Imports System.Text.Json

Public Class Form1

    Private ReadOnly s_specialKnownTimeZones As New Dictionary(Of String, String) From {
            {"Argentina Standard Time", "Argentina Standard Time"},
            {"Central European Summer Time", "Central European Daylight Time"},
            {"Eastern European Summer Time", "E. Europe Daylight Time"}
       }

    Friend Const MmolLUnitsDivisor As Single = 18
    Friend s_timeWithMinuteFormat As String
    Friend s_timeWithoutMinuteFormat As String
    Public Const MilitaryTimeWithMinuteFormat As String = "HH:mm"
    Public Const MilitaryTimeWithoutMinuteFormat As String = "HH:mm"
    Public Const TwelveHourTimeWithMinuteFormat As String = "h:mm tt"
    Public Const TwelveHourTimeWithoutMinuteFormat As String = "h:mm tt"

    Public ReadOnly s_unitsStrings As New Dictionary(Of String, String) From {
                {"MG_DL", "mg/dl"},
                {"MGDL", "mg/dl"},
                {"MMOL_L", "mmol/L"},
                {"MMOLL", "mmol/L"}
            }

    Friend Property BgUnits As String
    Friend Property BgUnitsString As String
    Public HomePageBasalRow As Integer
    Public HomePageInsulinRow As Integer
    Public HomePageMealRow As Integer
    Public s_clientTimeZone As TimeZoneInfo
    Public s_clientTimeZoneName As String
    Public s_criticalLow As Integer
    Public s_limitHigh As Integer
    Public s_limitLow As Integer
    Public s_timeZoneList As List(Of TimeZoneInfo)
    Public s_useLocalTimeZone As Boolean
    Public scalingNeeded As Boolean

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Dim testDataWithPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SampleUserData.json")
        Using document As JsonDocument = JsonDocument.Parse(File.ReadAllText(testDataWithPath))
            Dim node As JsonElement = document.RootElement
            s_timeZoneList = TimeZoneInfo.GetSystemTimeZones.ToList

            Try
                Select Case ShapeOf node
                    Case {"bgUnits": bgUnits As String,
                          "timeFormat": timeFormat As String,
                          "clientTimeZoneName": clientTimeZoneName As String,
                          "averageSGFloat": averageSGFloat As Single,
                          "sgs": [
                                    {"sg": sg As Integer,
                                     "datetime": datetime As Date,
                                     "timeChange": timeChange As Boolean,
                                     "sensorState": sensorState As String,
                                     "kind": kind As String,
                                     "version": version As Integer,
                                     "relativeOffset": relativeOffset As Integer
                                    }
                                 ]
                         }
                        s_clientTimeZoneName = clientTimeZoneName
                        If Not s_unitsStrings.TryGetValue(bgUnits, BgUnitsString) Then
                            If averageSGFloat > 40 Then
                                BgUnitsString = "mg/dl"
                                Me.BgUnits = "MG_DL"
                            Else
                                BgUnitsString = "mmol/L"
                                Me.BgUnits = "MMOL_L"
                            End If
                        End If
                        If BgUnitsString = "mg/dl" Then
                            scalingNeeded = False
                            HomePageBasalRow = 400
                            HomePageInsulinRow = 342
                            HomePageMealRow = 50
                            s_criticalLow = 50
                            s_limitHigh = 180
                            s_limitLow = 70
                        Else
                            scalingNeeded = True
                            HomePageBasalRow = 22
                            HomePageInsulinRow = 19
                            HomePageMealRow = CSng(Math.Round(50 / MmolLUnitsDivisor, 0, MidpointRounding.ToZero))
                            s_criticalLow = HomePageMealRow
                            s_limitHigh = 10
                            s_limitLow = 3.9
                        End If
                        If s_useLocalTimeZone Then
                            s_clientTimeZone = TimeZoneInfo.Local
                        Else
                            s_clientTimeZone = CalculateTimeZone(clientTimeZoneName)
                        End If
                        s_timeWithMinuteFormat = If(timeFormat = "HR_12", TwelveHourTimeWithMinuteFormat, MilitaryTimeWithMinuteFormat)
                        s_timeWithoutMinuteFormat = If(timeFormat = "HR_12", TwelveHourTimeWithoutMinuteFormat, MilitaryTimeWithoutMinuteFormat)
                End Select
            Catch ex As Exception
                Stop
                'Throw
            End Try

        End Using
        Stop
    End Sub

    Friend Function CalculateTimeZone(clientTimeZoneName As String) As TimeZoneInfo
        If My.Settings.UseLocalTimeZone Then
            Return TimeZoneInfo.Local
        End If
        If clientTimeZoneName = "NaN" Then
            Return Nothing
        End If
        Dim clientTimeZone As TimeZoneInfo
        Dim id As String = ""
        If Not s_specialKnownTimeZones.TryGetValue(clientTimeZoneName, id) Then
            id = clientTimeZoneName
        End If

        If id.Contains("Daylight") Then
            clientTimeZone = s_timeZoneList.Where(Function(t As TimeZoneInfo)
                                                      Return t.DaylightName = id
                                                  End Function).FirstOrDefault
            If clientTimeZone IsNot Nothing Then
                Return clientTimeZone
            End If
        End If

        clientTimeZone = s_timeZoneList.Where(Function(t As TimeZoneInfo)
                                                  Return t.StandardName = id
                                              End Function).FirstOrDefault
        If clientTimeZone IsNot Nothing Then
            Return clientTimeZone
        End If

        Return s_timeZoneList.Where(Function(t As TimeZoneInfo)
                                        Return t.DisplayName = id
                                    End Function).FirstOrDefault
    End Function

End Class
