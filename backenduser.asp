<%
	class backenduser
		public i_idUser, s_username, s_password, i_admin, i_idstore, s_email, i_iddept, i_idsupervisor, s_warehouse

		private conn

		private sub Class_Initialize()
			set conn = Server.CreateObject("ADODB.Connection")
			conn.Open pDatabaseConnectionString
			call resetProperties
		end sub

		public default function Init(user)
			if isNumeric(user) then
				idUser = user
			else
				username = user
			end if

			set Init = Me
		end function

		' idUser Begin
			public property let idUser(iIdUser)
				if isNumeric(iIdUser) AND NOT IsEmpty(iIdUser) AND NOT IsNull(iIdUser) then
					set rs = Server.CreateObject("ADODB.RecordSet")
						strsql =	"SELECT * " & vbcrlf &_
									"FROM dbo.backenduser AS bu WITH(NOLOCK) " & vbcrlf &_
									"WHERE idUser = " & iIdUser
						rs.Open strsql, conn

						if not rs.eof then
							i_idUser       = rs("idUser")
							s_username     = rs("username")
							s_password     = rs("password")
							i_admin        = rs("admin")
							i_idstore      = rs("idstore")
							s_email        = rs("email")
							i_iddept       = rs("iddept")
							i_idsupervisor = rs("idsupervisor")
							s_warehouse    = rs("warehouse")
						else
							call resetProperties
						end if
						rs.close
					set rs = nothing
				else
					call resetProperties
				end if
			end property
			public property get idUser()
				idUser = i_idUser
			end property

			public function getidUser()
				getidUser = idUser
			end function
		' idUser End

		' username Begin
			public property let username(sUsername)
				if NOT IsEmpty(sUsername) AND NOT IsNull(sUsername) then
					set rs = Server.CreateObject("ADODB.RecordSet")
						strsql =	"SELECT * " & vbcrlf &_
									"FROM dbo.backenduser AS bu WITH(NOLOCK) " & vbcrlf &_
									"WHERE username = '" & sUsername & "'"
						rs.Open strsql, conn

						if not rs.eof then
							i_idUser       = rs("idUser")
							s_username     = rs("username")
							s_password     = rs("password")
							i_admin        = rs("admin")
							i_idstore      = rs("idstore")
							s_email        = rs("email")
							i_iddept       = rs("iddept")
							i_idsupervisor = rs("idsupervisor")
							s_warehouse    = rs("warehouse")
						else
							call resetProperties
						end if
						rs.close
					set rs = nothing
				else
					call resetProperties
				end if
			end property
			public property get username()
				username = s_username
			end property

			public function getUsername()
				getUsername = username
			end function
		' username End

		' Gets Begin
			public property get password()
				password = s_password
			end property

			public property get admin()
				admin = i_admin
			end property

			public property get idStore()
				idStore = i_idstore
			end property

			public property get email()
				email = s_email
			end property

			public property get idDept()
				idDept = i_iddept
			end property

			public property get idsupervisor()
				idsupervisor = i_idsupervisor
			end property

			public property get warehouse()
				warehouse = s_warehouse
			end property
		' Gets End

		' Get functions Begin
			public function getPassword()
				getPassword = password
			end function

			public function getAdmin()
				getAdmin = admin
			end function

			public function getidStore()
				getidStore = idStore
			end function

			public function getEmail()
				getEmail = email
			end function

			public function getidDept()
				getidDept = idDept
			end function

			public function getidsupervisor()
				getidsupervisor = idsupervisor
			end function

			public function getWarehouse()
				getWarehouse = warehouse
			end function
		' Get functions End

		public function updatePassword(sPassword)
			if i_idUser <> 0 then
				strsql =	"UPDATE dbo.backenduser SET password = '" & sPassword & "' " & vbcrlf &_
							"WHERE i_idUser = " & i_idUser
				conn.Execute strsql
				setNewPassword = true
			else
				setNewPassword = false
			end if
		end function

		public function updateEmail(sEmail)
			if i_idUser <> 0 then
				strsql =	"UPDATE dbo.backenduser SET password = '" & sPassword & "' " & vbcrlf &_
							"WHERE i_idUser = " & i_idUser
				conn.Execute strsql
				setNewPassword = true
			else
				setNewPassword = false
			end if
		end function

		private function isLoaded()
			if i_idUser = 0 then
				isLoaded = false
				exit function
			end if
		end function

		private function update(field, value)
			if NOT isLoaded() OR isNull(value) then
				update = false
				exit function
			end if

			select case lcase(field)
				case "username"
					strsql =	"UPDATE dbo.backenduser SET username = '" & field & "' " & vbcrlf &_
								"WHERE i_idUser = " & i_idUser
				case "password"
					strsql =	"UPDATE dbo.backenduser SET password = '" & field & "' " & vbcrlf &_
								"WHERE i_idUser = " & i_idUser
				case "admin"
					strsql =	"UPDATE dbo.backenduser SET admin = " & field & " " & vbcrlf &_
								"WHERE i_idUser = " & i_idUser
				case "idstore"
					strsql =	"UPDATE dbo.backenduser SET idstore = " & field & " " & vbcrlf &_
								"WHERE i_idUser = " & i_idUser
				case "email"
					strsql =	"UPDATE dbo.backenduser SET email = '" & field & "' " & vbcrlf &_
								"WHERE i_idUser = " & i_idUser
				case "iddept"
					strsql =	"UPDATE dbo.backenduser SET iddept = " & field & " " & vbcrlf &_
								"WHERE i_idUser = " & i_idUser
				case "idsupervisor"
					strsql =	"UPDATE dbo.backenduser SET idsupervisor = " & field & " " & vbcrlf &_
								"WHERE i_idUser = " & i_idUser
				case "warehouse"
					strsql =	"UPDATE dbo.backenduser SET warehouse = '" & field & "' " & vbcrlf &_
								"WHERE i_idUser = " & i_idUser
			end select

			conn.Execute strsql

			update = true
		end function

		private sub resetProperties()
			i_idUser       = 0
			s_username     = ""
			s_password     = ""
			i_admin        = 0
			i_idstore      = 0
			s_email        = ""
			i_iddept       = 0
			i_idsupervisor = 0
			s_warehouse    = ""
		end sub

		private sub Class_Terminate()
			conn.close
			set conn = nothing
		end sub
	end class
%>