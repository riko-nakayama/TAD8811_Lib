Imports System
Imports System.Management
Imports System.IO.Ports
Imports System.Text.RegularExpressions

''' <summary>
''' TAD8811とUSB通信を行うためのクラスライブラリです。
''' USBデバイスはVirtual COM DeviceとしてWindowsに認識されています。
''' </summary>
Public Module TAD8811_Usb

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
	Private Function WriteFile(ByVal tx_buf As Byte(), ByVal len As Integer) As Boolean

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
	Private Function ReadFile(ByVal rx_buf As Byte(), ByVal rx_len As Integer, ByRef rx_num As Integer) As Boolean

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

End Module

