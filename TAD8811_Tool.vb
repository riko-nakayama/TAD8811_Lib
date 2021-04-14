Imports System


Public Module TAD8811_Tool


	Private OldTime(15) As Integer
	Private NewTime(15) As Integer
    Public Const MAIN_TIME As Integer = 0
    Public Const USB_TIME As Integer = 1

	Public Sub SetNow(ByVal t_idx As Integer)
		OldTime(t_idx) = Date.Now.Second
	End Sub

	Public Function IsTimerLimit(ByVal t_idx As Integer, ByVal wait_sec As Integer) As Boolean

		If wait_sec >= 60 Then
			wait_sec = 59
		End If

		NewTime(t_idx) = Date.Now.Second

		If NewTime(t_idx) < OldTime(t_idx) Then
			NewTime(t_idx) += 60
		End If

		If NewTime(t_idx) - OldTime(t_idx) > wait_sec Then
			Return True
		Else
			Return False
		End If

	End Function

	Public Sub SetNowMiliSecond(ByVal t_idx As Integer)
		OldTime(t_idx) = Date.Now.Millisecond
	End Sub

	Public Function IsTimerLimitMiliSecond(ByVal t_idx As Integer, ByVal wait_msec As Integer) As Boolean

		NewTime(t_idx) = Date.Now.Millisecond

		If NewTime(t_idx) < OldTime(t_idx) Then
			NewTime(t_idx) += 1000
		End If

		If NewTime(t_idx) - OldTime(t_idx) > wait_msec Then
			Return True
		Else
			Return False
		End If

	End Function

	Public Function AsciiToDec(ByVal data As Byte) As Byte

		Dim dec As Byte = 0

		If data < &H3A AndAlso data > &H2F Then
			dec = CByte(data - &H30 And &HF)
		Else
			dec = &HFF
		End If

		Return dec

	End Function

	Public Function HexToAscii(ByVal data As Byte) As Byte

		Dim asc As Byte = 0

		If data < 10 Then
			asc = CByte(data + &H30)
		Else
			asc = CByte(data + &H37)
		End If

		Return asc

	End Function

End Module

