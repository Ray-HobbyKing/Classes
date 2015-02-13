<!-- #include file = "backenduser.asp" -->
<%
	'----------------------------------------------------------------------------------------'
	' @File         : class_knowledge.asp                                                    '
	' @Author       : Raymond.Lam                                                            '
	' @Created Date : 2014-11-26                                                             '
	' @Revision     :                                                                        '
	'----------------------------------------------------------------------------------------'

	class knowledgeBase
		public intIdKnowledge, strTitle, strContent, strDescription, dateCreatedDate, intIdUser, dateLastModified, intLastModifiedUser
		public conn, rs, debug
		public oBkuser


		private sub Class_Initialize()
			set conn = Server.CreateObject("ADODB.Connection")
			conn.Open pDatabaseConnectionString

			set oBkuser = (new backenduser)(session("bkusername"))
			debug = 1
		end sub


		private sub Class_Terminate()
			set bkuser = nothing
			conn.close
			set conn = nothing
		end sub


		public default function Init(seed)
			if isNumeric(seed) then
				setIdKnowledge(seed)
			else
				setKnowledge(seed)
			end if

			set Init = Me
		end function


		private function addNewChecklist()

		end function


		public function tableFormat()

		end function


		public sub resetProperties()
			intIdKnowledge      = 0
			strTitle            = ""
			strContent          = ""
			strDescription      = ""
			dateCreatedDate     = ""
			intIdUser           = 0
			dateLastModified    = ""
			intLastModifiedUser = 0
		end sub

		' idKnowledge Begin
			public property let idKnowledge(i_idKnowledge)
				' Nothing to do here
			end property
			public property get idKnowledge()
				idKnowledge = intIdKnowledge
			end property

			private function setIdKnowledge(i_idKnowledge)
				if isNumeric(i_idKnowledge) AND NOT IsEmpty(i_idKnowledge) AND NOT IsNull(i_idKnowledge) then
					set rs = Server.CreateObject("ADODB.RecordSet")
						strsql =	"SELECT * " & vbcrlf &_
									"FROM dbo.backendKnowledgeBase AS kb WITH(NOLOCK) " & vbcrlf &_
									"WHERE kb.idKnowledge = " & i_idKnowledge
						call rsSQL(rs, strsql)

						if not rs.eof then
							intIdKnowledge      = rs("idKnowledge")
							strTitle            = rs("title")
							strContent          = rs("content")
							strDescription      = rs("description")
							dateCreatedDate     = rs("createdDate")
							intIdUser           = rs("creator")
							dateLastModified    = rs("lastModified")
							intLastModifiedUser = rs("lastModifiedUser")
						else
							call resetProperties
						end if
						rs.close
					set rs = nothing
				else
					call resetProperties
				end if
			end function
		' idKnowledge End

		' title Begin
			public property let title(s_title)
				strsql =	"UPDATE dbo.backendKnowledgeBase SET title = '" & s_title & "' "
							"WHERE idKnowledge = " & idKnowledge
				call connSQL(strsql)
				setIdKnowledge(idKnowledge)
			end property
			public property get title()
				title = strTitle
			end property
		' title End

		' content Begin
			public property let content(s_content)
				strsql =	"UPDATE dbo.backendKnowledgeBase SET content = '" & s_content & "' " &_
							"WHERE idKnowledge = " & idKnowledge
				call connSQL(strsql)
				setIdKnowledge(idKnowledge)
			end property
			public property get content()
				content = strContent
			end property
		' content End

		' description Begin
			public property let description(s_description)
				strsql =	"UPDATE dbo.backendKnowledgeBase SET description = '" & s_description & "' " &_
							"WHERE idKnowledge = " & idKnowledge
				call connSQL(strsql)
				setIdKnowledge(idKnowledge)
			end property
			public property get description()
				description = strDescription
			end property
		' description End

		' createdDate Begin
			private property let createdDate(d_createdDate)
				if isEmpty(d_createdDate) OR isNull(d_createdDate) OR NOT isDate(d_createdDate) OR d_createdDate = "now" then
					d_createdDate = "getDate()"
				else
					d_createdDate = "'" & d_createdDate & "'"
				end if
				strsql =	"UPDATE dbo.backendKnowledgeBase SET createdDate = " & d_createdDate & " " &_
							"WHERE idKnowledge = " & idKnowledge
				call connSQL(strsql)
				setIdKnowledge(idKnowledge)
			end property
			public property get createdDate()
				createdDate = d_createdDate
			end property
		' createdDate End

		' creator Begin
			private property let creator(i_creator)
				strsql =	"UPDATE dbo.backendKnowledgeBase SET creator = " & oBkuser.idUser & " " &_
							"WHERE idKnowledge = " & idKnowledge
				call connSQL(strsql)
				setIdKnowledge(idKnowledge)
			end property
			public property get creator()
				creator = intIdUser
			end property
		' creator End

		' lastModified Begin
			public property let lastModified(d_lastModified)
				strsql =	"UPDATE dbo.backendKnowledgeBase SET lastModified = getDate() " &_
							"WHERE idKnowledge = " & idKnowledge
				call connSQL(strsql)
				setIdKnowledge(idKnowledge)
			end property
			public property get lastModified()
				lastModified = d_lastModified
			end property
		' lastModified End

		' lastModifiedUser Begin
			private property let lastModifiedUser(s_content)
				strsql =	"UPDATE dbo.backendKnowledgeBase SET lastModifiedUser = " & oBkuser.idUser & " " &_
							"WHERE idKnowledge = " & idKnowledge
				call connSQL(strsql)
				setIdKnowledge(idKnowledge)
			end property
			public property get lastModifiedUser()
				lastModifiedUser = intLastModifiedUser
			end property
		' lastModifiedUser End


		public function addKnowledge(objDict)
			for each key in objDict.keys
				fields = fields & key & ", "
				if isNumeric(objDict(key)) then
					values = values & objDict(key) & ", "
				else
					values = values & "'" & objDict(key) & "', "
				end if
			next

			on error resume next
				strsql =	"INSERT INTO dbo.backendKnowledgeBase (" & left(fields, len(fields) -1) & ") VALUES " &_
							"(" & left(values, len(values) -1) & ")"
				call connSQL(strsql)
				addKnowledge = true
				exit function
			on error goto 0
			addKnowledge = false
		end function

		private sub connSQL(strsql)
			if debug then response.write "<pre>" & strsql & "</pre>"
			conn.Execute strsql
		end sub

		private sub rsSQL(rs, strsql)
			if debug then response.write "<pre>" & strsql & "</pre>"
			rs.Open strsql, conn
		end sub

		private sub nextRecord()
			rs.movenext
		end sub

		private sub previousRecord()
			rs.moveprevious
		end sub

		private function iif(con, rTrue, rFalse)
			iif = rFalse
			if con then iif = rTrue
		end function

	end class
%>