Imports pfcls

Module Module1

    Public Sub Main()
        Dim run As New setSurfaceFinishToDeco
        run.setParameter()
    End Sub

    Public Class setSurfaceFinishToDeco

        Public Sub setParameter()
            Dim asyncConnection As IpfcAsyncConnection = Nothing
            Dim model As IpfcModel
            Dim activeserver As IpfcServer
            Dim paramown As IpfcParameterOwner
            Dim ipparam As IpfcParameter
            Dim ipbaseparam As IpfcBaseParameter
            Dim paramval As IpfcParamValue
            Dim Paraname As String = "SURFACE_FINISH"

            Dim session As IpfcBaseSession
            Dim Moditem As CMpfcModelItem

            Try
                asyncConnection = (New CCpfcAsyncConnection).Connect(Nothing, Nothing, Nothing, Nothing)
                session = asyncConnection.Session
                activeserver = session.GetActiveServer
                model = session.CurrentModel


                If model Is Nothing Then
                    MsgBox("Model is not present",, "Script message")
                    Environment.Exit(0)
                End If

                If (Not model.Type = EpfcModelType.EpfcMDL_PART) And (Not model.Type = EpfcModelType.EpfcMDL_ASSEMBLY) Then
                    MsgBox("Model is not a solid",, "Script message")
                    Environment.Exit(0)
                End If

                If Not activeserver.IsObjectCheckedOut(activeserver.ActiveWorkspace, model.FileName) Then
                    MsgBox("Please check out model first...",, "Script Message")
                    Environment.Exit(0)
                End If

                paramown = model
                ipparam = paramown.GetParam(Paraname)
                ipbaseparam = ipparam

                Moditem = New CMpfcModelItem

                paramval = Moditem.CreateStringParamValue(CStr("-"))
                ipbaseparam.Value = paramval

                MsgBox("Surface finish has been set to '-'",, "Script Message")
                asyncConnection.Disconnect(1)

            Catch ex As Exception
            End Try
        End Sub

    End Class

End Module


