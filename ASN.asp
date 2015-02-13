<%
	class ASN
		private conn

		private idASN, createdDate, createdBy, ASNtype, referenceNum, idsupplier, receivingWarehouse, active, dateCompleted

		private sub Class_Initialize()
			set conn = Server.CreateObject("ADODB.Connection")
			conn.Open pDatabaseConnectionString
			call resetProperties
		end sub

		private sub Class_Terminate()
			conn.close
			set conn = nothing
		end sub

		private sub resetProperties()
			
		end sub

	end class
%>