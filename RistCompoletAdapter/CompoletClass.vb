Imports System.Reflection
Imports System.IO

<ComClass(CompoletClass.ClassId, CompoletClass.InterfaceId, CompoletClass.EventsId)> _
Public Class CompoletClass

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "62cc624f-2bf1-4a9f-a522-8df10827232e"
    Public Const InterfaceId As String = "693e4af5-0188-4c5f-9dff-668b27ffff23"
    Public Const EventsId As String = "1cb7f676-3927-4717-b1b5-91a4eb86e25d"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub
    Public NJCompolet As OMRON.Compolet.CIP.NJCompolet
    Private c_njVariables As New NjVariables
    Private c_resqVariables As New ResQVariables
    Private c_portLocal As String
    Private c_IpLocal As String
    Private c_ComLogPath As String = "C:\Comlog1.txt"

    Public Function Connect(ByVal ip As String, ByVal port As Integer) As Boolean
        c_portLocal = CStr(port)
        c_IpLocal = ip
        Return ConnectNJ()
    End Function

    Public Function SendLotInfoToPlc(ByVal strRecipe As String, ByVal strMCNo As String, ByVal strLotNo As String, ByVal strPackage As String, ByVal strProcess As String, ByVal strOPE As String, ByVal strDevice As String) As Boolean
        If IsConnect() = False Then
            ProductLog("SendLotInfoToPlc : Can not connect")
            Return False
        End If

        c_resqVariables.Recipe = strRecipe '"Unknown"
        c_resqVariables.Eqp_no = strMCNo 'msvb.Right(My.Settings.MCNo, My.Settings.MCNo.Length - 3)
        c_resqVariables.Lot_no = strLotNo 'ListData(0).WaferFabLotNo
        c_resqVariables.Package = strPackage 'ListData(0).PackageName
        c_resqVariables.Process = strProcess '"Dicing"
        c_resqVariables.Ope_name = strOPE '"Dicer"
        c_resqVariables.Device = strDevice 'ListData(0).DeviceName
        Dim njResult As Boolean = False
        Try
            NJCompolet.WriteVariable("NJ_RECIPE", c_resqVariables.Recipe)
            NJCompolet.WriteVariable("NJ_EQP_NO", c_resqVariables.Eqp_no)
            NJCompolet.WriteVariable("NJ_LOT_NO", c_resqVariables.Lot_no)
            NJCompolet.WriteVariable("NJ_PACKAGE", c_resqVariables.Package)
            NJCompolet.WriteVariable("NJ_PROCESS", c_resqVariables.Process)
            NJCompolet.WriteVariable("NJ_OPE_NAME", c_resqVariables.Ope_name)
            NJCompolet.WriteVariable("NJ_DEVICE", c_resqVariables.Device)
            njResult = True
        Catch ex As Exception
            njResult = False
            ProductLog("SendLotInfoToPlc :" & "NJ_RECIPE" & c_resqVariables.Recipe & _
                                            "NJ_EQP_NO" & c_resqVariables.Eqp_no & _
                                            "NJ_LOT_NO" & c_resqVariables.Lot_no & _
                                            "NJ_PACKAGE" & c_resqVariables.Package & _
                                            "NJ_PROCESS" & c_resqVariables.Process & _
                                            "NJ_OPE_NAME" & c_resqVariables.Ope_name & _
                                            "NJ_DEVICE" & c_resqVariables.Device & _
                           vbNewLine + MethodBase.GetCurrentMethod().Name + vbNewLine + ex.ToString())
            Return njResult
        End Try

        Dim strDataProduct As String
        strDataProduct = "SendLotInfoToPlc >RecipeName:" & strRecipe & ",MCN0:" & strMCNo & ",LotNO:" & strLotNo & ",Package:" & strPackage & ",Process :" & strProcess & ",OPE:" & strOPE & ",Device:" & strDevice
        ProductLog(strDataProduct)

        Return njResult

    End Function

    Private Function ConnectNJ() As Boolean
        'NjCompolet1
        Try
            NJCompolet = New OMRON.Compolet.CIP.NJCompolet
            NJCompolet.UseRoutePath = False
            NJCompolet.PeerAddress = c_IpLocal
            NJCompolet.LocalPort = Int16.Parse(c_portLocal)
            NJCompolet.Active = True
            If NJCompolet.RunStatus <> -1 Then
                Return True
            End If
            Return False
        Catch er As Exception
            ProductLog("ConnectNJ :" & MethodBase.GetCurrentMethod().Name & " : " & er.Message.ToString)
            Return False
        End Try
    End Function

    Public Function SendEndCleaning() As Boolean
        If IsConnect() = False Then
            ProductLog("SendEndCleaning : Can not connect")
            Return False
        End If
        Try
            NJCompolet.WriteVariable("NJ_PROCESS", "Cleaning")
        Catch ex As Exception
            ProductLog("IoTEndCleaning :" & "NJ_PROCESS Cleaning" & vbNewLine + MethodBase.GetCurrentMethod().Name + vbNewLine + ex.ToString())
            Return False
        End Try

        Dim strDataProduct As String
        strDataProduct = "IoTEndCleaning >Process: Cleaning"
        ProductLog(strDataProduct)
        Return True
    End Function

    Private Sub ProductLog(ByVal strData As String)
        Try
            Dim pathData As String = c_ComLogPath

            If File.Exists(c_ComLogPath) = True Then
                Dim FileData As New FileInfo(pathData)
                If FileData.Length > 2000000 Then
                    File.Delete(pathData)
                End If
            End If

            Using sw As New StreamWriter(pathData, True)
                sw.WriteLine(Format(Now, "yyyy/MM/dd HH:mm:ss") & " " & strData)
            End Using

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub
    Sub SetDirectoryPath(ByVal comlogPath As String)
        c_ComLogPath = comlogPath
    End Sub

    Public Function IsConnect() As Boolean
        Try
            Return NJCompolet.IsConnected
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function HeartBeat() As Integer
        Try
            Return HeartBeat
        Catch ex As Exception

        End Try
    End Function

    Public Class NjVariables
        Public NJRequest As String
        Public Mean As String
        Public Max As String
        Public Min As String
        Public SD As String
        Public Range As String
        Public MeanTag As String
        Public MaxTag As String
        Public MinTag As String
        Public SDTag As String
        Public RangeTag As String
        Public NJRequestTag As String
    End Class

    Class ResQVariables
        Public Time As DateTime
        Public Start_time As DateTime
        Public Item_no As String
        Public Batch_id As String
        Public Data_id As String
        Public Mesure_item As String
        Public Mesure_point As Integer
        Public Data As String

        Public Recipe As String
        Public Eqp_no As String
        Public Lot_no As String
        Public Package As String
        Public Process As String
        Public Ope_name As String
        Public Device As String

        Public L_param1 As String
        Public D_param1 As String

        'Public RecipeTag As String
        'Public Eqp_noTag As String
        'Public Lot_NoTag As String
        'Public PackageTag As String
        'Public ProcessTag As String
        'Public Ope_nameTag As String
        'Public DeviceTag As String


    End Class

End Class


