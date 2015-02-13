<%
	class fileUploader
		Const ForReading = 1, ForWriting = 2, ForAppending = 3
		Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

		public OverwriteFiles, SetMaxSize, strFileName, boolOverwriteFiles, intSetMaxSize, fullFileName
		public error
		public upload

		private sub Class_Initialize()
			set conn = Server.CreateObject("ADODB.Connection")
			conn.ConnectionTimeout = 820
			conn.CommandTimeout = 120
			conn.Open pDatabaseConnectionString

			set FSO = Server.CreateObject("Scripting.FileSystemObject")
		end sub

		private sub Class_Terminate()
			set bkuser = nothing
			conn.close
			set conn = nothing
		end sub


		public default function Init(upload)
			upload.OverwriteFiles = false
			upload.SetMaxSize 50000000, True

			set Init = Me
		end function


		public property let SaveFileName(s_File)
			strFileName = s_File
		end property
		public property get SaveFileName()
			SaveFileName = strFileName
		end property
		private function isFileNameSet()
			isFileNameSet = false
			if SaveFileName = "" then isFileNameSet = true
		end function


		public property let OverwriteFiles(bool)
			if isNumeric(bool) then
				if bool = 0 then
					upload.OverwriteFiles = false
				else
					upload.OverwriteFiles = true
					bool = 1
				end if
				boolOverwriteFiles = bool
			end if
		end property
		public property get OverwriteFiles()
			OverwriteFiles = boolOverwriteFiles
		end property



		public property let SetMaxSize(size)
			if isNumeric(size) then
				upload.SetMaxSize size, true
				intSetMaxSize = size
			end if
		end property
		public property get SetMaxSize()
			SetMaxSize = intSetMaxSize
		end property


		public function uploadCredentials()
			if isFileNameSet then
				upload.LogonUser Const_backend_upload_login_domain, Const_backend_upload_login_userID, Const_backend_upload_login_PW
				upload.Save Const_backend_upload_path

				fullFileName = Const_backend_upload_path & "purchasing\" & SaveFileName
			end if
		end function


		public function checkCSV()
			strFileName = File.FileName

			strExt = right(strFileName, Len(strFileName) - Instrrev(strFileName, "."))
			select case strExt
				case "csv"
					file.SaveAs fullFileName
				case else
					response.write "Please check your file extension, only csv is accepted. Click <a href=""shipping_containerUpload.asp"">here</a> to go back"
					response.end
			end select
		end function

	end class
%>