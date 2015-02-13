<%
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' tableBuilder is a class to build a fast formatted table with recordset
	'
	' @package     tableBuilder
	' @version     0.1
	' @author      Raymond@HobbyKing
	' @email       raymond.lam@hobbyking.com
	' @license     GPL
	' @copyright   2015 HobbyKing
	' @link        http://www.hobbyking.com
	'
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	class tableBuilder
		private tb_rs
		private arr_headers, replaceHeader
		private wholeTable

		private str_class, int_isSortable, bool_collapse

		private sub Class_Initialize()
			bool_collapse  = " collapse"
			int_isSortable = 0
		end sub

		public default function Init(rs)
			set tb_rs = rs

			set Init = Me
		end function

		private sub Class_Terminate()
			set tb_rs = nothing
		end sub

		public sub headerOrder(strHeaders)
			if strHeaders <> "" then
				arr_headers = split(strHeaders, ",")
			end if
		end sub


		public property let custom_column(d)

		end property


		private function render_header()
			' if headerOrder is called, i.e. user enters the order of showing data
			if vartype(arr_headers) > vbArray then
				' arr_headers has value
				if ubound(arr_headers) > 0 then
					notExist = ""

					' check if the user input headers exist
					for each header in arr_headers
						if not fieldExists(header) then
							notExist = header & ","
						end if
					next

					' return error on user input header doesn't exist
					if len(notExist) > 0 then
						notExist = left(notExist, len(notExist) - 1)
						render_header = "Error: " & notExist & " not exist(s)"
						exit function
					end if

					'response.write "replaceHeader: " & replaceHeader & "<br />"
					arr_replace = split(left(replaceHeader, len(replaceHeader) - 1), ";")

					for each header in arr_headers

						if ubound(arr_replace) > 0 then
							for each replacement in arr_replace
								a = split(replacement, "|||")
								if a(0) = header then
									header = a(1)
								end if
							next
						end if
						render_header = render_header & "<th>" & header & "</th>"
					next

					render_header = "<thead><tr>" & render_header & "</tr></thead>"
				end if
			else
				arr_replace = split(replaceHeader, ";")

				for each field in tb_rs.fields
					if ubound(arr_replace) > 0 then
						for each replacement in arr_replace
							a = split(replacement, "|||")
							if a(0) = header then
								header = a(1)
							end if
						next
					else
						header = field.name
					end if
					render_header = render_header & "<th>" & header & "</th>"
				next

				render_header = "<thead><tr>" & render_header & "</tr></thead>"
			end if
		end function

		public sub relabelHeader(strHeader, strValue)
			if trim(strHeader) <> "" AND trim(strValue) <> "" then
				replaceHeader = replaceHeader & strHeader & "|||" & strValue & ";"
			end if
		end sub


		private function fieldExists(seed)
			for each field in tb_rs.fields
				if lcase(seed) = lcase(field.name) then
					fieldExists = true
					exit function
				end if
			next
			fieldExists = false
		end function

		private function columnWidth()
			columnWidth = len(tb_rs.fields)
		end function


		public property let tableClasses(strClass)
			if len(trim(strClass)) > 0 then
				str_class = trim(strClass)
			end if
		end property
		public property get tableClasses()
			tableClasses = " class=""" & str_class & borderCollapse & """"
		end property


		public property let tableSortable(flag)
			if flag = true OR flag = 1 then
				int_isSortable = flag
			end if
		end property
		public property get tableSortable()
			tableSortable = ""
			if int_isSortable then
				tableSortable = " id=""tableToSort"""
			end if
		end property


		public property let borderCollapse(flag)
			if flag = false OR flag = 0 then
				bool_collapse = ""
			end if
		end property
		public property get borderCollapse()
			borderCollapse = bool_collapse
		end property


		public function render_as_table()
			headers = render_header()
			' When user has entered the wanted fields
			if ubound(arr_headers) > 0 then

				if left(headers, 6) = "Error:" then
					render_as_table = headers
					exit function
				end if

				if not tb_rs.eof then
					while not tb_rs.eof
						body = body & "<tr>"
						for each header in arr_headers
							body = body & "<td>" & rs(header) & "</td>"
						next
						body = body & "</tr>"
						tb_rs.movenext
					wend
				end if

				render_as_table = "<table" & tableSortable & tableClasses & " border=""1"" cellpadding=""5"" cellspacing=""0"">" & headers & "<tbody>" & body & "</tbody></table>"
			else
				' Generating full set of field data
				if not tb_rs.eof then
					while not tb_rs.eof
						row = tb_rs.getrows
						tb_rs.movenext
					wend
				end if
				render_as_table = "<table" & tableSortable & tableClasses & " border=""1"" cellpadding=""5"" cellspacing=""0"">" & headers & "<tbody>" & body & "</tbody></table>"
			end if
		end function
	end class
%>