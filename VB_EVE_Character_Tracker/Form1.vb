Imports System.Net
Imports System.Xml
Imports System.IO
Imports System.DateTime
Imports System.TimeZone

Public Class Form1

    Dim timezone As TimeZone = CurrentTimeZone
    Dim intOffset As Integer
    Dim listTypeID As New List(Of Integer)
    Dim listTypeName As New List(Of String)
    Dim web31W As WebClient = New WebClient()
    Dim web31T As WebClient = New WebClient()
    Dim web32W As WebClient = New WebClient()
    Dim web32T As WebClient = New WebClient()
    Dim webFW As WebClient = New WebClient()
    Dim webFT As WebClient = New WebClient()

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        intOffset = CInt(timezone.GetUtcOffset(DateTime.Now).TotalHours())

        QueryServer()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Public Sub QueryServer()

        Dim shrtError As Short = 0
        Dim strTraining31 As New Uri("https://api.eveonline.com/char/SkillInTraining.xml.aspx?keyID=2793180&vCode=HHOqJgmPKpMLmZjcIhPPjSQXF6A04XVVutSWQtjTcF84VXTemPdUyhDFKOd3dF4p&characterID=91810030")
        Dim strWallet31 As New Uri("https://api.eveonline.com/char/AccountBalance.xml.aspx?keyID=2793180&vCode=HHOqJgmPKpMLmZjcIhPPjSQXF6A04XVVutSWQtjTcF84VXTemPdUyhDFKOd3dF4p&characterID=91810030")
        Dim strTraining32 As New Uri("https://api.eveonline.com/char/SkillInTraining.xml.aspx?keyID=3033583&vCode=arLeZxr9AUImT3WGVOtOsMh3hTTjhginXFpu429WIcohRRKQPtJddSHZEnDdNmxq&characterID=91995291")
        Dim strWallet32 As New Uri("https://api.eveonline.com/char/AccountBalance.xml.aspx?keyID=3033583&vCode=arLeZxr9AUImT3WGVOtOsMh3hTTjhginXFpu429WIcohRRKQPtJddSHZEnDdNmxq&characterID=91995291")
        Dim strTrainingFThis As New Uri("https://api.eveonline.com/char/SkillInTraining.xml.aspx?keyID=3033615&vCode=Cz8RmzKhGbTNqGxlc55ERETVQyRIs9tLsABySu1G7irkVPFjsErR8IOaYffsBxTw&characterID=1042504230")
        Dim strWalletFThis As New Uri("https://api.eveonline.com/char/AccountBalance.xml.aspx?keyID=3033615&vCode=Cz8RmzKhGbTNqGxlc55ERETVQyRIs9tLsABySu1G7irkVPFjsErR8IOaYffsBxTw&characterID=1042504230")

        AddHandler web31W.DownloadStringCompleted, AddressOf Wallet31
        AddHandler web31T.DownloadStringCompleted, AddressOf Training31
        AddHandler web32W.DownloadStringCompleted, AddressOf Wallet32
        AddHandler web32T.DownloadStringCompleted, AddressOf Training32
        AddHandler webFW.DownloadStringCompleted, AddressOf WalletF
        AddHandler webFT.DownloadStringCompleted, AddressOf TrainingF

        Try
            web31W.DownloadStringAsync(strWallet31)
            web31T.DownloadStringAsync(strTraining31)
            web32W.DownloadStringAsync(strWallet32)
            web32T.DownloadStringAsync(strTraining32)
            webFW.DownloadStringAsync(strWalletFThis)
            webFT.DownloadStringAsync(strTrainingFThis)
        Catch ex As Exception
            txbStatus.Text = "Can't get data for some reason."
            shrtError = 1
        End Try

        If shrtError = 0 Then
            txbStatus.Text = "Last attempt: " + Now().ToString("MMM-d 'at' h:mm:ss tt")

        Else
            txbStatus.Text = "Error reaching server."
        End If
    End Sub

    Public Sub Wallet31(ByVal sender As Object, ByVal e As DownloadStringCompletedEventArgs)
        RemoveHandler web31W.DownloadStringCompleted, AddressOf Wallet31
        Dim xmlData As XElement
        Dim strBalance As Decimal

        Try
            xmlData = XElement.Parse(e.Result)
            strBalance = CDec(xmlData...<row>.@balance.ToString)
            txbWallet31.Text = strBalance.ToString("N2")
        Catch ex As Exception
            txbStatus.Text = "Exception: Endpoint could not be reached."
            txbWallet31.Text = "N/A"
        End Try
    End Sub

    Public Sub Wallet32(ByVal sender As Object, ByVal e As DownloadStringCompletedEventArgs)
        RemoveHandler web32W.DownloadStringCompleted, AddressOf Wallet32
        Dim xmlData As XElement
        Dim strBalance As Decimal

        Try
            xmlData = XElement.Parse(e.Result)
            strBalance = CDec(xmlData...<row>.@balance.ToString)
            txbWallet32.Text = strBalance.ToString("N2")
        Catch ex As Exception
            txbStatus.Text = "Exception: Endpoint could not be reached."
            txbWallet32.Text = "N/A"
        End Try
    End Sub

    Public Sub WalletF(ByVal sender As Object, ByVal e As DownloadStringCompletedEventArgs)
        RemoveHandler webFW.DownloadStringCompleted, AddressOf WalletF
        Dim xmlData As XElement
        Dim strBalance As Decimal

        Try
            xmlData = XElement.Parse(e.Result)
            strBalance = CDec(xmlData...<row>.@balance.ToString)
            txbWalletFThis.Text = strBalance.ToString("N2")
        Catch ex As Exception
            txbStatus.Text = "Exception: Endpoint could not be reached."
            txbWalletFThis.Text = "N/A"
        End Try
    End Sub

    Public Sub Training31(ByVal sender As Object, ByVal e As DownloadStringCompletedEventArgs)
        RemoveHandler web31T.DownloadStringCompleted, AddressOf Training31
        Dim xmlData As XElement
        Dim endTime As String = ""
        Dim strSkillTypeID As Integer
        Dim strSkillName As String
        Dim strSkillLevel As String
        Try
            xmlData = XElement.Parse(e.Result)
            endTime = xmlData...<trainingEndTime>.Value
            endTime = DateAdd(DateInterval.Hour, intOffset, CDate(endTime))
            txbTraining31.Text = CDate(endTime).ToString("ddd MMM d, yyyy 'at' h:mm:ss tt")

            strSkillTypeID = CInt(xmlData...<trainingTypeID>.Value)
            strSkillLevel = xmlData...<trainingToLevel>.Value
            strSkillName = GetSkillName(strSkillTypeID)

            txbSkill31.Text = strSkillName & ", Rank " & strSkillLevel
        Catch ex As Exception
            txbSkill31.Text = "Skill queue empty."
            txbTraining31.Text = "N/A"
        End Try
    End Sub

    Public Sub Training32(ByVal sender As Object, ByVal e As DownloadStringCompletedEventArgs)
        RemoveHandler web32T.DownloadStringCompleted, AddressOf Training32
        Dim xmlData As XElement
        Dim endTime As String = ""
        Dim strSkillTypeID As Integer
        Dim strSkillName As String
        Dim strSkillLevel As String

        Try
            xmlData = XElement.Parse(e.Result)
            endTime = xmlData...<trainingEndTime>.Value
            endTime = DateAdd(DateInterval.Hour, intOffset, CDate(endTime))
            txbTraining32.Text = CDate(endTime).ToString("ddd MMM d, yyyy 'at' h:mm:ss tt")
            
            strSkillTypeID = CInt(xmlData...<trainingTypeID>.Value)
            strSkillLevel = xmlData...<trainingToLevel>.Value
            strSkillName = GetSkillName(strSkillTypeID)

            txbSkill32.Text = strSkillName & ", Rank " & strSkillLevel
        Catch ex As Exception
            txbSkill32.Text = "Skill queue empty."
            txbTraining32.Text = "N/A"
        End Try
    End Sub

    Public Sub TrainingF(ByVal sender As Object, ByVal e As DownloadStringCompletedEventArgs)
        RemoveHandler webFT.DownloadStringCompleted, AddressOf TrainingF
        Dim xmlData As XElement
        Dim endTime As String = ""
        Dim strSkillTypeID As Integer
        Dim strSkillName As String
        Dim strSkillLevel As String
        Dim currentTime = Date.Now().ToUniversalTime(), cachedUntil As Date
        Dim dateDifference As TimeSpan
        Dim msResult As Integer = 0

        Try
            xmlData = XElement.Parse(e.Result)
            endTime = xmlData...<trainingEndTime>.Value
            endTime = DateAdd(DateInterval.Hour, intOffset, CDate(endTime))
            cachedUntil = xmlData...<cachedUntil>.Value
            dateDifference = cachedUntil - currentTime
            cachedUntil = DateAdd(DateInterval.Hour, intOffset, CDate(cachedUntil))
            txbNextPull.Text = "Next attempt: " & DateAdd(DateInterval.Second, 10, CDate(cachedUntil)).ToString("MMM-d 'at' h:mm:ss tt")
            msResult = dateDifference.TotalMilliseconds
            TimerTraining.Enabled = False
            TimerTraining.Interval = msResult + 10000
            TimerTraining.Enabled = True

            txbTrainingFThis.Text = CDate(endTime).ToString("ddd MMM d, yyyy 'at' h:mm:ss tt")

            strSkillTypeID = CInt(xmlData...<trainingTypeID>.Value)
            strSkillLevel = xmlData...<trainingToLevel>.Value
            strSkillName = GetSkillName(strSkillTypeID)

            txbSkillFThis.Text = strSkillName & ", Rank " & strSkillLevel
        Catch ex As Exception
            txbSkillFThis.Text = "Skill queue empty."
            txbTrainingFThis.Text = "N/A"
        End Try
    End Sub

    Public Function GetSkillName(ByVal TypeID As Integer) As String
        Dim index As Integer = BinarySearch(TypeID)

        If Not index = -1 Then
            Return listTypeName.Item(index)
        End If

        Return "N/A"
    End Function

    Public Sub LoadTypeIDs()
        Try
            Dim strArray(1) As String
            Dim TypeList As New List(Of String)
            TypeList = My.Resources.TypeIDs.Kronos_98534_typeID.Split(vbNewLine).ToList()

            For counter As Integer = 0 To TypeList.Count - 1
                strArray = TypeList.Item(counter).Split(",")
                listTypeID.Add(CInt(strArray(0)))
                listTypeName.Add(strArray(1))
            Next
        Catch ex As Exception
        End Try

    End Sub

    Public Function BinarySearch(ByVal TypeID As Integer) As Integer
        Dim left As Integer = 0
        Dim right As Integer = listTypeID.Count()
        Dim mid As Integer = (left + right) / 2
        Dim currentTypeID As Integer
        Try
            While left <= right
                currentTypeID = listTypeID.Item(mid)
                If currentTypeID > TypeID Then
                    right = mid - 1
                    mid = (left + right) / 2
                ElseIf currentTypeID < TypeID Then
                    left = mid + 1
                    mid = (left + right) / 2
                ElseIf currentTypeID = TypeID Then
                    Return mid
                End If
            End While
        Catch ex As Exception

        End Try

        Return -1
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadTypeIDs()
        web31W.Proxy() = Nothing
        web31T.Proxy() = Nothing
        web32W.Proxy() = Nothing
        web32T.Proxy() = Nothing
        webFW.Proxy() = Nothing
        webFT.Proxy() = Nothing
    End Sub

    Private Sub TimerTraining_Tick(sender As Object, e As EventArgs) Handles TimerTraining.Tick
        QueryServer()
    End Sub
End Class
