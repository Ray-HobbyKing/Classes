<%
	class shippingNumbers
		private conn
		public i_idShip, s_shipmentRef, s_createdDate, s_exportedDate, s_uploadedDate, i_shippedAll, i_active

		public i_debug

		'--------------------------------------------------------------------------------------------------------'
		' Class Basic Begin                                                                                      '
		'--------------------------------------------------------------------------------------------------------'
			private sub Class_Initialize()
				set conn = Server.CreateObject("ADODB.Connection")
				conn.Open pDatabaseConnectionString
				call resetProperties
				debug = 0
			end sub

			private sub Class_Terminate()
				conn.close
				set conn = nothing
			end sub

			public default function Init(shipment)
				if isNumeric(shipment) then
					setidShip(shipment)
				else
					setShipmentRef(shipment)
				end if

				set Init = Me
			end function

			public sub resetProperties()
				i_idShip       = 0
				s_shipmentRef  = ""
				s_createdDate  = ""
				s_exportedDate = ""
				s_uploadedDate = ""
				i_shippedAll   = -1
				i_active       = -1
			end sub

			private sub connSQL(strsql)
				if debug then response.write "<pre>" & strsql & "</pre>"
				conn.Execute strsql
			end sub

			private sub rsSQL(rs, strsql)
				if debug then response.write "<pre>" & strsql & "</pre>"
				rs.Open strsql, conn
			end sub
		'--------------------------------------------------------------------------------------------------------'
		' Class Basic End                                                                                        '
		'--------------------------------------------------------------------------------------------------------'

		'--------------------------------------------------------------------------------------------------------'
		' Debugging Begin                                                                                        '
		'--------------------------------------------------------------------------------------------------------'
			public property let debug(idebug)
				i_debug = idebug
			end property

			public property get debug()
				debug = i_debug
			end property
		'--------------------------------------------------------------------------------------------------------'
		' Debugging End                                                                                          '
		'--------------------------------------------------------------------------------------------------------'

		'--------------------------------------------------------------------------------------------------------'
		' idShip Begin                                                                                           '
		'--------------------------------------------------------------------------------------------------------'
			public property let idShip(iidShip)
				' Primary key, won't do let / update
			end property

			public property get idShip()
				idShip = i_idShip
			end property

			private function setidShip(iidShip)
				if isNumeric(iidShip) AND NOT IsEmpty(iidShip) AND NOT IsNull(iidShip) then
					set rs = Server.CreateObject("ADODB.RecordSet")
						strsql =	"SELECT * " & vbcrlf &_
									"FROM dbo.shippingNumbers AS sn WITH(NOLOCK) " & vbcrlf &_
									"WHERE sn.idShip = " & iidShip
						call rsSQL(rs, strsql)

						if not rs.eof then
							i_idShip       = rs("idShip")
							s_shipmentRef  = rs("shipmentRef")
							s_createdDate  = rs("createdDate")
							s_exportedDate = rs("exportedDate")
							s_uploadedDate = rs("uploadedDate")
							i_shippedAll   = rs("shippedAll")
							i_active       = rs("active")
						else
							call resetProperties
						end if
						rs.close
					set rs = nothing
				else
					call resetProperties
				end if
			end function

			public function getidShip()
				getidShip = idShip
			end function
		'--------------------------------------------------------------------------------------------------------'
		' idShip end                                                                                             '
		'--------------------------------------------------------------------------------------------------------'

		'--------------------------------------------------------------------------------------------------------'
		' Shipment Reference Begin                                                                               '
		'--------------------------------------------------------------------------------------------------------'
			public property let shipmentRef(sshipmentRef)
				if idShip <> 0 then
					strsql =	"UPDATE dbo.shippingNumbers SET shipmentRef = '" & sshipmentRef & "' " & vbcrlf &_
								"WHERE idShip = " & idShip
					call connSQL(strsql)
					s_shipmentRef = sshipmentRef
				end if
			end property

			public property get shipmentRef()
				shipmentRef = s_shipmentRef
			end property

			private function setShipmentRef(sshipmentRef)
				if isNumeric(sshipmentRef) AND NOT IsEmpty(sshipmentRef) AND NOT IsNull(sshipmentRef) then
					set rs = Server.CreateObject("ADODB.RecordSet")
						strsql =	"SELECT * " & vbcrlf &_
									"FROM dbo.shippingNumbers AS sn WITH(NOLOCK) " & vbcrlf &_
									"WHERE sn.shipmentRef = " & sshipmentRef
						call rsSQL(rs, strsql)

						if not rs.eof then
							i_idShip       = rs("idShip")
							s_shipmentRef  = rs("shipmentRef")
							s_createdDate  = rs("createdDate")
							s_exportedDate = rs("exportedDate")
							s_uploadedDate = rs("uploadedDate")
							i_shippedAll   = rs("shippedAll")
							i_active       = rs("active")
						else
							call resetProperties
						end if
						rs.close
					set rs = nothing
				else
					call resetProperties
				end if
			end function

			public function getShipmentRef()
				getShipmentRef = shipmentRef
			end function
		'--------------------------------------------------------------------------------------------------------'
		' Shipment Reference End                                                                                 '
		'--------------------------------------------------------------------------------------------------------'

		'--------------------------------------------------------------------------------------------------------'
		' Created Date Begin                                                                                     '
		'--------------------------------------------------------------------------------------------------------'
			public property let createDate(sCreateDate)
				if lcase(sCreateDate) = "getdate()" then
					strCreatedDate = "getDate()"
				else
					strCreatedDate = "'" & sCreateDate & "'"
				end if
				strsql =	"UPDATE dbo.shippingNumbers SET createdDate = " & strCreatedDate & " " & vbcrlf &_
							"WHERE idShip = " & idShip
				call connSQL(strsql)
				setidShip(idShip)
			end property

			public property get createDate()
				createdDate = s_createdDate
			end property

			public function getCreatedDate()
				getCreatedDate = createdDate
			end function
		'--------------------------------------------------------------------------------------------------------'
		' Created Date End                                                                                       '
		'--------------------------------------------------------------------------------------------------------'

		'--------------------------------------------------------------------------------------------------------'
		' Exported Date Begin                                                                                    '
		'--------------------------------------------------------------------------------------------------------'
			public property let exportedDate(sexportedDate)
				if lcase(sexportedDate) = "getdate()" then
					strExportedDate = "getDate()"
				else
					strExportedDate = "'" & sexportedDate & "'"
				end if
				strsql =	"UPDATE dbo.shippingNumbers SET createdDate = " & strExportedDate & " " & vbcrlf &_
							"WHERE idShip = " & idShip
				call connSQL(strsql)
				setidShip(idShip)
			end property

			public property get exportedDate()
				exportedDate = s_exportedDate
			end property

			public function getExportedDate()
				getCreatedDate = createdDate
			end function
		'--------------------------------------------------------------------------------------------------------'
		' Exported Date End                                                                                      '
		'--------------------------------------------------------------------------------------------------------'

		'--------------------------------------------------------------------------------------------------------'
		' Uploaded Date Begin                                                                                    '
		'--------------------------------------------------------------------------------------------------------'
			public property let uploadedDate(suploadedDate)
				if lcase(suploadedDate) = "getdate()" then
					strUploadedDate = "getDate()"
				else
					strUploadedDate = "'" & suploadedDate & "'"
				end if
				strsql =	"UPDATE dbo.shippingNumbers SET uploadedDate = " & strUploadedDate & " " & vbcrlf &_
							"WHERE idShip = " & idShip
				call connSQL(strsql)
				setidShip(idShip)
			end property

			public property get uploadedDate()
				uploadedDate = s_exportedDate
			end property

			public function getUploadDate()
				getUploadDate = uploadedDate
			end function
		'--------------------------------------------------------------------------------------------------------'
		' Uploaded Date End                                                                                      '
		'--------------------------------------------------------------------------------------------------------'

		'--------------------------------------------------------------------------------------------------------'
		' Shipped All Begin                                                                                      '
		'--------------------------------------------------------------------------------------------------------'
			public property let shippedAll(iShippedAll)
				strsql =	"UPDATE dbo.shippingNumbers SET shippedAll = " & iShippedAll & " " & vbcrlf &_
							"WHERE idShip = " & idShip
				call connSQL(strsql)
				setidShip(idShip)
			end property

			public property get shippedAll()
				shippedAll = i_shippedAll
			end property

			public function getShippedAll()
				getShippedAll = shippedAll
			end function
		'--------------------------------------------------------------------------------------------------------'
		' Shipped All End                                                                                        '
		'--------------------------------------------------------------------------------------------------------'

		'--------------------------------------------------------------------------------------------------------'
		' active Begin                                                                                           '
		'--------------------------------------------------------------------------------------------------------'
			public property let active(iActive)
				strsql =	"UPDATE dbo.shippingNumbers SET active = " & iActive & " " & vbcrlf &_
							"WHERE idShip = " & idShip
				call connSQL(strsql)
				setidShip(idShip)
			end property

			public property get active()
				active = i_active
			end property

			public function getActive()
				getActive = active
			end function
		'--------------------------------------------------------------------------------------------------------'
		' active End                                                                                             '
		'--------------------------------------------------------------------------------------------------------'
	end class
%>