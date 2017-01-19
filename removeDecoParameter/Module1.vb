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

                model = session.CurrentModel


                If model Is Nothing Then
                    Throw New Exception("Model not present")
                End If

                paramown = model
                ipparam = paramown.GetParam(Paraname)
                ipbaseparam = ipparam

                Moditem = New CMpfcModelItem

                paramval = Moditem.CreateStringParamValue(CStr("-"))
                ipbaseparam.Value = paramval

                MsgBox("Surface finish has been set to '-'")
                asyncConnection.Disconnect(1)

            Catch ex As Exception
            End Try
        End Sub

    End Class

End Module


