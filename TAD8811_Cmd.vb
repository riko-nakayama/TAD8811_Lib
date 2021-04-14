Imports System
Imports System.Threading
Imports System.IO.Ports
Imports System.Management
Imports System.Text.RegularExpressions

''' <summary>
''' TAD8811へ指令を発行するグローバルモジュールです。
''' </summary>
Public Module TAD8811_Cmd

	Private clsUsb As TAD8811_Usb = New TAD8811_Usb()

	Private Const STX As Byte = &H2
	Private Const ETX As Byte = &H3
	Private Const ACK As Byte = &H6
	Private Const NAK As Byte = &H15
	Private Const REQ_LEN As Integer = 64
	Private Const RES_SVALL_LEN As Integer = 1032
	Private Const RES_SVRD_LEN As Integer = 252
	Private Const RES_SVWR_LEN As Integer = 252

	Private UsbTxBuf(63) As Byte
	Private UsbRxBuf(16383) As Byte
	Private UsbRsBuf(16383) As Byte

	''' <summary>
	''' USBデバイス接続＋COMポート接続状態を参照します。
	''' </summary>
	''' <returns>USBデバイス接続＋COMポート接続ならTrue、いずれから未接続ならFalseを応答します。</returns>
	Public ReadOnly Property IsUsbConnect As Boolean

		Get
			Return clsUsb.IsUsbConnect
		End Get

	End Property

	''' <summary>
	''' TAD8811がUSB接続済みならCOMポートをOpenします。
	''' </summary>
	''' <returns>COMポートを正常にOpenできたらTrue、失敗したらFalseを応答します。</returns>
	Public Function OpenDevice() As Boolean

		Return clsUsb.OpenDevice()

	End Function

	''' <summary>
	''' TAD8811が接続済みのCOMポートをCloseします。
	''' </summary>
	''' <returns>COMポートを正常にCloseできたらTrue、失敗したらFalseを応答します。</returns>
	Public Function CloseDevice() As Boolean

		Return clsUsb.CloseDevice()

	End Function

	''' <summary>
	''' <para>TAD8811のパラメータを初期化します。</para>
	''' <para>ID69（コントロールスイッチ）：サーボオフ時の偏差リセット有効＋プロファイル動作完了時のフラグ自動クリア有効</para>
	''' <para>ID209（アラームマスク）：多回転アラーム検出禁止</para>
	''' </summary>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function InitTAD8811() As Boolean

		If Not WriteParam(69, &H3) Then         ' ID69 コントロールスイッチ　サーボオフ時の偏差リセット有効　プロファイル動作完了時のフラグ自動クリア有効
			Return False
		End If

		If Not WriteParam(209, 8) Then          ' ID209 アラームマスク　多回転アラーム検出禁止
			Return False
		End If

		Return True

	End Function

	''' <summary>
	''' フィードバックデータ（位置・速度・電流）を取得します。
	''' </summary>
	''' <param name="pos">現在位置を格納する参照変数</param>
	''' <param name="vel">現在速度を格納する参照変数</param>
	''' <param name="cur">現在電流を格納する参照変数</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function GetFeedback(ByRef pos As Integer, ByRef vel As Integer, ByRef cur As Integer) As Boolean

		If Not TAD8811_Cmd.ReadParam(40, pos) Then
			Return False
		End If

		If Not TAD8811_Cmd.ReadParam(41, vel) Then
			Return False
		End If

		If Not TAD8811_Cmd.ReadParam(42, cur) Then
			Return False
		End If

		Return True

	End Function

	''' <summary>
	''' サーボオン指令を発行します。
	''' </summary>
	''' <param name="ctrl_mode">TAD8811の制御モードを指定します。</param>
	''' <param name="wait_msec">サーボオン後の待ち時間を設定します。（既定値=100msec）</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function ServoOn(ByVal ctrl_mode As Integer, Optional wait_msec As Integer = 100) As Boolean

		If Not WriteParam(31, ctrl_mode) Then
			Return False
		End If

		If Not WriteParam(30, 1) Then
			Return False
		End If

		Thread.Sleep(wait_msec)

		Return True

	End Function

	''' <summary>
	''' サーボオフ指令を発行します。
	''' </summary>
	''' <param name="wait_msec">サーボオフ後の待ち時間を設定します。（既定値=100msec）</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function ServoOff(Optional wait_msec As Integer = 100) As Boolean

		If Not WriteParam(30, 0) Then
			Return False
		End If

		Thread.Sleep(wait_msec)

		Return True

	End Function

	''' <summary>
	''' TAD8811にポジションリセット指令を発行します。
	''' </summary>
	''' <param name="reset_pos">ポジションリセット後の現在値を設定します。</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function PositionReset(Optional reset_pos As Integer = 0) As Boolean

		Dim cmd As Integer

		If Not WriteParam(39, reset_pos) Then       ' ポジションリセット値を設定
			Return False
		End If

		If Not ReadParam(30, cmd) Then
			Return False
		End If

		cmd = cmd Or &H4000

		If Not WriteParam(30, cmd) Then
			Return False
		End If

		Thread.Sleep(500)

		cmd = cmd And (Not &H4000)

		If Not WriteParam(30, cmd) Then
			Return False
		End If

		Return True

	End Function

	''' <summary>
	''' 正回転方向を設定します。
	''' </summary>
	''' <param name="dir">0ならモータセンサ側から見てCW方向を正回転方向、1ならCCW方向を正回転方向に設定します。</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function SetDirection(ByVal dir As Integer) As Boolean
		'dir = 0 既定値 dir = 1 反転
		If Not WriteParam(72, dir) Then
			Return False
		End If

		Return True

	End Function

	''' <summary>
	''' 位置決め動作を開始します。
	''' </summary>
	''' <param name="pos">目標位置（絶対位置）を設定します。</param>
	''' <param name="vel">位置決め速度[rpm]を設定します。</param>
	''' <param name="acc">加速度[10rpm/s]を設定します。</param>
	''' <param name="dcc">減速度[10rpm/s]を設定します。</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function StartProfile(ByVal pos As Integer, ByVal vel As Integer, ByVal acc As Integer, ByVal dcc As Integer) As Boolean

		Dim cmd As Integer = 0

		If Not WriteParam(32, pos) Then
			Return False
		End If

		If Not WriteParam(33, vel) Then
			Return False
		End If

		If Not WriteParam(34, acc) Then
			Return False
		End If

		If Not WriteParam(35, dcc) Then
			Return False
		End If

		If Not ReadParam(30, cmd) Then
			Return False
		End If

		cmd = cmd And Not &H30              ' BIT4 ハードストップOFF　BIT5 スムーズストップOFF
		cmd = cmd Or &H82                   ' BIT1 プロファイル動作許可ON　BIT7 加減速有効ON

		If Not WriteParam(30, cmd) Then     ' プロファイル動作開始
			Return False
		End If

		Return True

	End Function

	''' <summary>
	''' 一定速送り（速度制御）を開始します。
	''' </summary>
	''' <param name="vel">速度[rpm]を設定します。</param>
	''' <param name="acc">加速度[10rpm/s]を設定します。</param>
	''' <param name="dcc">減速度[10rpm/s]を設定します。</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function StartJog(ByVal vel As Integer, ByVal acc As Integer, ByVal dcc As Integer) As Boolean

		Dim cmd As Integer = 0

		If Not WriteParam(34, acc) Then
			Return False
		End If

		If Not WriteParam(35, dcc) Then
			Return False
		End If

		If Not ReadParam(30, cmd) Then
			Return False
		End If

		cmd = cmd And Not &H30              ' BIT4 ハードストップOFF　BIT5 スムーズストップOFF
		cmd = cmd Or &H80                   ' BIT7 加減速有効ON

		If Not WriteParam(30, cmd) Then
			Return False
		End If

		If Not WriteParam(37, vel) Then     ' ID37 リアルタイム速度指令更新でJog開始
			Return False
		End If

		Return True

	End Function

	''' <summary>
	''' 位置決め完了（インポジション到達）まで待機します。
	''' </summary>
	''' <param name="wait_msec">位置決め完了後の待ち時間[msec]を設定します。</param>
	''' <param name="timeout_sec">位置決め完了タイムアウト時間[sec]を設定します。</param>
	''' <param name="wait_judge">位置決め完了判定待ち時間[msec]を設定します。既定値は100msecです。</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function WaitInpositon(ByVal wait_msec As Integer, ByVal timeout_sec As Integer, Optional wait_judge As Integer = 100) As Boolean

		Dim sts As Integer = 0

		Thread.Sleep(wait_judge)

		TAD8811_Tool.SetNow(0)

		Do

			If TAD8811_Tool.IsTimerLimit(0, timeout_sec) Then

				Return False

			End If

			ReadParam(20, sts)

		Loop While (sts And &H2) = &H2

		Thread.Sleep(wait_msec)

		Return True

	End Function

	''' <summary>
	''' 原点復帰完了（原点到達）まで待機します。
	''' </summary>
	''' <param name="wait_msec">原点復帰完了後の待ち時間[msec]を設定します。</param>
	''' <param name="timeout_sec">原点復帰完了タイムアウト時間[sec]を設定します。</param>
	''' <param name="wait_judge">原点付記完了判定待ち時間[msec]を設定します。既定値は100msecです。</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function WaitHoming(ByVal wait_msec As Integer, ByVal timeout_sec As Integer, Optional wait_judge As Integer = 200) As Boolean

		Dim sts As Integer = 0

		Thread.Sleep(wait_judge)

		TAD8811_Tool.SetNow(1)

		Do

			If TAD8811_Tool.IsTimerLimit(1, timeout_sec) = True Then

				Return False

			End If

			ReadParam(20, sts)

		Loop While (sts And &H400) = &H400

		Thread.Sleep(wait_msec)

		Return True

	End Function

	''' <summary>
	''' 即停止します。
	''' </summary>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function HardStop() As Boolean

		Dim cmd As Integer = 0

		If Not ReadParam(30, cmd) Then
			Return False
		End If

		cmd = cmd And Not &H30              ' BIT4 ハードストップOFF　BIT5 スムーズストップOFF
		cmd = cmd Or &H10                   ' BIT4 ハードストップON

		If Not WriteParam(30, cmd) Then
			Return False
		End If

		Return True

	End Function

	''' <summary>
	''' 減速停止します。
	''' </summary>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function SmoothStop() As Boolean

		Dim cmd As Integer = 0

		If Not ReadParam(30, cmd) Then
			Return False
		End If

		cmd = cmd And Not &H30              ' BIT4 ハードストップOFF　BIT5 スムーズストップOFF
		cmd = cmd Or &H20                   ' BIT5 スムーズストップON

		If Not WriteParam(30, cmd) Then
			Return False
		End If

		Return True

	End Function

	''' <summary>
	''' アラーム解除します。
	''' </summary>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function AlarmRest() As Boolean

		Dim cmd As Integer = 0

		If Not ReadParam(30, cmd) Then
			Return False
		End If

		cmd = cmd Or &H8                    ' BIT3 アラームリセットON

		If Not WriteParam(30, cmd) Then
			Return False
		End If

		Thread.Sleep(500)

		cmd = cmd And Not &H8               ' BIT3 アラームリセットOFF

		If Not WriteParam(30, cmd) Then
			Return False
		End If

		Return True

	End Function

	''' <summary>
	''' 全てのパラメータを読み込みます。一度に256個のパラメータを読み込みます。
	''' </summary>
	''' <param name="page">パラメータのページ番号を設定します。</param>
	''' <param name="data">パラメータを格納する参照変数（配列）を設定します。</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function ReadParamAll(ByVal page As Integer, ByRef data As Integer()) As Boolean

		Try
			UsbTxBuf(0) = STX
			UsbTxBuf(1) = Convert.ToByte("s"c)
			UsbTxBuf(2) = CByte(page)
			UsbTxBuf(3) = 0

			Dim sum As Integer = 0

			For i = 1 To 62 - 1

				If i >= 4 Then
					UsbTxBuf(i) = 0
				End If

				sum += UsbTxBuf(i)
			Next

			UsbTxBuf(62) = ETX
			UsbTxBuf(63) = Convert.ToByte(sum)

			If Not clsUsb.WriteFile(UsbTxBuf, REQ_LEN) Then
				Return False
			End If

			If Not CheckUsbRead(RES_SVALL_LEN) Then
				Return False
			End If

			Dim a As Integer = 0
			Dim b As Integer = 0

			While a < 256
				data(a) = Convert.ToInt32(UsbRxBuf(b + 4))
				data(a) = data(a) Or (Convert.ToInt32(UsbRxBuf(b + 5)) << 8 And &HFF00)
				data(a) = data(a) Or (Convert.ToInt32(UsbRxBuf(b + 6)) << 16 And &HFF0000)
				data(a) = data(a) Or (Convert.ToInt32(UsbRxBuf(b + 7)) << 24 And &HFF000000)
				a += 1
				b += 4
			End While

			Return True

		Catch ex As Exception

			Return False

		End Try

	End Function

	''' <summary>
	''' パラメータを読み込みます。1個のパラメータを読み込みます。
	''' </summary>
	''' <param name="id">パラメータID番号を設定します。</param>
	''' <param name="data">パラメータを格納する参照変数を設定します。</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function ReadParam(ByVal id As Integer, ByRef data As Integer) As Boolean

		Try

			UsbTxBuf(0) = STX
			UsbTxBuf(1) = Convert.ToByte("r"c)
			UsbTxBuf(2) = 0
			UsbTxBuf(3) = 0
			UsbTxBuf(4) = CByte(id And &HFF)
			UsbTxBuf(5) = CByte(id >> 8 And &HFF)
			UsbTxBuf(6) = CByte(id >> 16 And &HFF)
			UsbTxBuf(7) = CByte(id >> 24 And &HFF)

			Dim sum As Integer = 0

			For i = 1 To 62 - 1

				If i >= 8 Then
					UsbTxBuf(i) = 0
				End If

				sum += UsbTxBuf(i)
			Next

			UsbTxBuf(62) = ETX
			UsbTxBuf(63) = Convert.ToByte(sum And &HFF)

			If Not clsUsb.WriteFile(UsbTxBuf, REQ_LEN) Then
				Return False
			End If

			If Not CheckUsbRead(RES_SVRD_LEN) Then
				Return False
			End If

			data = UsbRxBuf(4)
			data = data Or (Convert.ToInt32(UsbRxBuf(5)) << 8 And &HFF00)
			data = data Or (Convert.ToInt32(UsbRxBuf(6)) << 16 And &HFF0000)
			data = data Or (Convert.ToInt32(UsbRxBuf(7)) << 24 And &HFF000000)

			Return True

		Catch ex As Exception

			Return False

		End Try

	End Function

	''' <summary>
	''' パラメータを変更します。
	''' </summary>
	''' <param name="id">パラメータID番号を設定します。</param>
	''' <param name="data">変更するデータを設定します。</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function WriteParam(ByVal id As Integer, ByVal data As Integer) As Boolean

		Try

			UsbTxBuf(0) = STX
			UsbTxBuf(1) = Convert.ToByte("w"c)
			UsbTxBuf(2) = 0
			UsbTxBuf(3) = 0
			UsbTxBuf(4) = CByte(id And &HFF)
			UsbTxBuf(5) = CByte(id >> 8 And &HFF)
			UsbTxBuf(6) = CByte(id >> 16 And &HFF)
			UsbTxBuf(7) = CByte(id >> 24 And &HFF)
			UsbTxBuf(8) = CByte(data And &HFF)
			UsbTxBuf(9) = CByte(data >> 8 And &HFF)
			UsbTxBuf(10) = CByte(data >> 16 And &HFF)
			UsbTxBuf(11) = CByte(data >> 24 And &HFF)

			Dim sum As Integer = 0

			For i = 1 To 62 - 1

				If i >= 12 Then
					UsbTxBuf(i) = 0
				End If

				sum += UsbTxBuf(i)
			Next

			UsbTxBuf(62) = ETX
			UsbTxBuf(63) = Convert.ToByte(sum And &HFF)

			If Not clsUsb.WriteFile(UsbTxBuf, REQ_LEN) Then
				Return False
			End If

			If id = 176 AndAlso data = 888 OrElse id = 176 AndAlso data = 999 Then
				Thread.Sleep(2000)
			End If

			If Not CheckUsbRead(RES_SVWR_LEN) Then
				Return False
			End If

			Return True

		Catch ex As Exception

			Return False

		End Try

	End Function

	''' <summary>
	''' リクエストコマンド（特殊指令）を発行します。
	''' </summary>
	''' <param name="cmd">リクエストコマンドコードを設定します。</param>
	''' <param name="buf">取得するデータを格納する参照変数（配列）を設定します。</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Public Function RequestCommandCode(ByVal cmd As Char, ByVal buf As Byte()) As Boolean

		Try

			UsbTxBuf(0) = STX
			UsbTxBuf(1) = Convert.ToByte(cmd)
			UsbTxBuf(2) = 0
			UsbTxBuf(3) = 0

			For i = 4 To 62 - 1
				UsbTxBuf(i) = buf(i)
			Next

			Dim sum As Integer = 0

			For i = 1 To 62 - 1
				sum += UsbTxBuf(i)
			Next

			UsbTxBuf(62) = ETX
			UsbTxBuf(63) = Convert.ToByte(sum And &HFF)

			If Not clsUsb.WriteFile(UsbTxBuf, REQ_LEN) Then
				Return False
			End If

			If cmd = "U"c Then
				Return True
			End If

			If cmd = "B"c Then
				Return True
			End If

			If Not CheckUsbRead(RES_SVRD_LEN) Then
				Return False
			End If

			Return True

		Catch ex As Exception

			Return False

		End Try

	End Function

	''' <summary>
	''' 取得したUSBデータをチェックします。
	''' </summary>
	''' <param name="rx_len">データ長さを設定します。</param>
	''' <returns>処理が正常終了した場合はTrue、異常終了した場合はFalseを応答します。</returns>
	Private Function CheckUsbRead(ByVal rx_len As Integer) As Boolean

		Try

			Dim rx_num As Integer = 0

			If Not clsUsb.ReadFile(UsbRxBuf, rx_len, rx_num) Then

				If Not clsUsb.WriteFile(UsbTxBuf, REQ_LEN) Then
					Return False
				End If

				If Not clsUsb.ReadFile(UsbRsBuf, rx_len, rx_num) Then
					Return False
				Else

					If rx_num >= rx_len Then

						For i = 0 To rx_len - 1
							UsbRxBuf(i) = UsbRsBuf(i + rx_num - rx_len)
						Next
					Else
						Return False
					End If
				End If
			End If

			If UsbRxBuf(0) <> ACK Then
				Return False
			End If

			Dim num As Integer = 0
			Dim sum As Byte = 0

			For i = 1 To rx_len - 2 - 1
				num += UsbRxBuf(i)
			Next

			sum = Convert.ToByte(num And &HFF)

			If sum <> UsbRxBuf(rx_len - 1) Then
				Return False
			End If

			Return True

		Catch ex As Exception

			Return False

		End Try

	End Function

	''' <summary>
	''' TAD8811とUSB通信を行うためのクラスライブラリです。
	''' USBデバイスはVirtual COM DeviceとしてWindowsに認識されています。
	''' </summary>
	Private Class TAD8811_Usb

		Private Const USB_BUF_SIZE As Integer = 8192
		Private Const EP1_FIFO_SIZE As Integer = 512
		Private Const EP2_FIFO_SIZE As Integer = 512

		Private UsbCom As SerialPort
		Private ComName As String
		Private _IsUsbConnect As Boolean
		Private _IsDeviceExist As Boolean

		''' <summary>
		''' TAD8811がUSB接続されているかWindowsデバイスマネージャーの情報を検索します。
		''' </summary>
		''' <returns>USB接続済みならTrue、未接続ならFalseを応答します。</returns>
		Private Function SearchDevice() As Boolean

			Try
				Dim check = New Regex("(COM[1-9][0-9]?[0-9]?)")
				Dim mcPnPEntity = New ManagementClass("Win32_PnPEntity")
				Dim manageObjCol = mcPnPEntity.GetInstances

				For Each manageObj As ManagementObject In manageObjCol

					Dim namePropertyValue = manageObj.GetPropertyValue("Name")

					If namePropertyValue Is Nothing Then
						Continue For
					End If

					Dim name = namePropertyValue.ToString

					If check.IsMatch(name) OrElse name.Contains("Tamagawa Seiki") Then

						If manageObj.Path.ToString.Contains("VID_0DC1") AndAlso (manageObj.Path.ToString.Contains("PID_0100") OrElse manageObj.Path.ToString.Contains("PID_0101")) Then

							Dim nameProperty = manageObj.GetPropertyValue("Name")
							Dim com = nameProperty.ToString
							Dim idx1 = com.IndexOf("(") + 1
							Dim idx2 = com.IndexOf(")")

							If idx2 > idx1 Then
								com = com.Substring(idx1, idx2 - idx1)
								ComName = com
								_IsDeviceExist = True
							Else
								_IsDeviceExist = False
							End If

							Return _IsDeviceExist
						End If
					End If
				Next

				Return False

			Catch

				Return False

			Finally
			End Try

		End Function

		''' <summary>
		''' TAD8811がUSB接続済みならCOMポートをOpenします。
		''' </summary>
		''' <returns>COMポートを正常にOpenできたらTrue、失敗したらFalseを応答します。</returns>
		Public Function OpenDevice() As Boolean

			Try

				If SearchDevice() = False Then

					_IsUsbConnect = False
					Return False

				End If

				If UsbCom Is Nothing Then

					UsbCom = New SerialPort(ComName, 115200, Parity.None, 8, StopBits.One)
					UsbCom.Handshake = Handshake.None
					UsbCom.ReadBufferSize = USB_BUF_SIZE
					UsbCom.ReadTimeout = 500
					UsbCom.WriteBufferSize = USB_BUF_SIZE
					UsbCom.WriteTimeout = 500
					UsbCom.DiscardNull = True
					UsbCom.Handshake = Handshake.XOnXOff
					UsbCom.ReceivedBytesThreshold = 1
					UsbCom.WriteBufferSize = 64
					UsbCom.ReadBufferSize = 64

				End If

				If UsbCom.IsOpen = False Then

					UsbCom.Open()

					UsbCom.DiscardInBuffer()
					UsbCom.DiscardOutBuffer()
					_IsUsbConnect = True

				End If

			Catch

				_IsUsbConnect = False

			End Try

			Return _IsUsbConnect

		End Function

		''' <summary>
		''' TAD8811が接続済みのCOMポートをCloseします。
		''' </summary>
		''' <returns>COMポートを正常にCloseできたらTrue、失敗したらFalseを応答します。</returns>
		Public Function CloseDevice() As Boolean

			Try

				_IsUsbConnect = False
				_IsDeviceExist = False
				UsbCom.Close()

				Return True

			Catch

				Return False

			End Try

		End Function

		''' <summary>
		''' USBデバイス接続＋COMポート接続状態を参照します。
		''' </summary>
		''' <returns>USBデバイス接続＋COMポート接続ならTrue、いずれから未接続ならFalseを応答します。</returns>
		Public ReadOnly Property IsUsbConnect As Boolean

			Get
				Return _IsUsbConnect
			End Get

		End Property

		''' <summary>
		''' USBデバイス接続状態を参照します。
		''' </summary>
		''' <returns>USBデバイス接続ならTrue、未接続ならFalseを応答します。</returns>
		Public ReadOnly Property IsDeviceExist As Boolean

			Get
				Return _IsDeviceExist
			End Get

		End Property

		''' <summary>
		''' USBデバイスにデータを転送します。
		''' </summary>
		''' <param name="tx_buf">送信データを格納する配列変数を設定します。</param>
		''' <param name="len">送信データ長さを設定します。</param>
		''' <returns>送信に成功したらTrue、失敗したらFalseを応答します。</returns>
		Public Function WriteFile(ByVal tx_buf As Byte(), ByVal len As Integer) As Boolean

			Try

				If UsbCom Is Nothing Then
					Return False
				End If

				If UsbCom.IsOpen = False Then
					Return False
				End If

				UsbCom.Write(tx_buf, 0, len)

				Return True

			Catch

				Debug.WriteLine("!!! WRITE USB ERROR CATCH !!!")
				Return False

			End Try

		End Function

		''' <summary>
		''' USBデバイスからデータを受信します。
		''' </summary>
		''' <param name="rx_buf">受信データを格納する配列変数を設定します。</param>
		''' <param name="rx_len">受信データ長さを設定します。</param>
		''' <param name="rx_num">実際に受信したデータ長さを格納する参照変数を設定します。</param>
		''' <returns>受信に成功したらTrue、失敗したらFalseを応答します。</returns>
		Public Function ReadFile(ByVal rx_buf As Byte(), ByVal rx_len As Integer, ByRef rx_num As Integer) As Boolean

			Try

				If UsbCom Is Nothing Then
					Return False
				End If

				If UsbCom.IsOpen = False Then
					Return False
				End If

				Dim off As Integer = 0
				Dim len As Integer = 0

				TAD8811_Tool.SetNowMiliSecond(0)

				Do
					len = UsbCom.BytesToRead

					If len <> 0 Then
						Debug.WriteLine(len.ToString)
					End If

					If len > 0 Then
						off += UsbCom.Read(rx_buf, off, len)
					End If

					If TAD8811_Tool.IsTimerLimitMiliSecond(0, 50) Then
						rx_num = off
						Debug.WriteLine("+++ USB READ TIME OUT +++")
						Return False
					End If

				Loop While off < rx_len

				rx_num = off

				Return True

			Catch

				Debug.WriteLine("!!! READ USB ERROR CATCH !!!")
				Return False

			End Try

		End Function

	End Class

End Module

