Imports System.Net
Imports System.Threading
Imports Newtonsoft.Json.Linq

<ComClass(weatherClass.ClassId, weatherClass.InterfaceId, weatherClass.EventsId)>
Public Class weatherClass

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "2252035c-414b-4040-bbcf-727b7b6faefb"
    Public Const InterfaceId As String = "d9dbc9ee-a2c8-4a16-82d1-5be73612111d"
    Public Const EventsId As String = "ae775dc9-a8e2-44fe-9bde-1066c7113bee"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Event Completed(ByVal result As String)
    Dim city As String
    Public Sub GetTempreture(cityName As String)

        Me.city = cityName
        Dim thread As Thread = New Thread(AddressOf GetAccountInfoFromApi)
        thread.Start()
    End Sub

    Private Sub GetAccountInfoFromApi()
        Dim accountInformationUrl As String = "http://api.openweathermap.org/data/2.5/weather?q=" & Me.city & "&appid=fa45908ebb85a1177ddac6c466c4b87b"
        Dim webClient As WebClient = New WebClient()
        'Return webClient.DownloadString(New Uri(accountInformationUrl))

        Try
            Dim result As String = webClient.DownloadString(New Uri(accountInformationUrl))
            Dim parsejson As JObject = JObject.Parse(result)
            RaiseEvent Completed(parsejson.SelectToken("main.temp").ToString())
        Catch ex As Exception

            'throw any other exception - this should not occur
            Throw
        End Try



    End Sub
End Class


