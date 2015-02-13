<!-- #include file = "class_barcodeTranslation.asp" -->
<%
	class inTransit
		public conn, debug

		public i_source, i_quantity, s_destination, i_destinationID

		public oProduct

		private sub Class_Initialize()
			set conn = Server.CreateObject("ADODB.Connection")
			conn.Open pDatabaseConnectionString
			debug = 1
			s_destination = ""
		end sub


		private sub Class_Terminate()
			conn.close
			set conn = nothing
		end sub


		public default function Init(seed)
			if isNumeric(seed) then
				' class_barcodeTranslation.asp
				set oProduct = (new barcodeTranslation)(seed)
			else
				response.write "<p>Barcode is not numeric</p>"
			end if
			set Init = Me
		end function


		public property let destination(countrycode)
			select case lcase(countrycode)
				case "au", "ar", "cn", "nl", "ru", "gb", "uk", "us"
					s_destination = countrycode
				case else
					response.write "<p>Destination is not set in one of the following: AU, AR, CN, NL, RU, GB, US</p>"
			end select
		end property
		public property get destination()
			destination = s_destination
		end property


		public property let transferQuantity(qty)
			if isNumeric(qty) then
				i_quantity = qty
			else
				response.write "<p>Quantity entered is not an integer</p>"
			end if
		end property
		public property get transferQuantity()
			transferQuantity = i_quantity
		end property


		private function getDestinationID()
			select case lcase(destination)
				case "au"
					i_destinationID = oProduct.au
				case "ar"
					i_destinationID = oProduct.ar
				case "cn"
					i_destinationID = oProduct.cn
				case "nl"
					i_destinationID = oProduct.nl
				case "ru"
					i_destinationID = oProduct.ru
				case "gb", "uk"
					i_destinationID = oProduct.uk
				case "us"
					i_destinationID = oProduct.us
			end select

			getDestinationID = i_destinationID
		end function

		public function confirmTransfer()
			destinationID = getDestinationID()
			transferQty   = transferQuantity()

			response.write "<p>Transfer " & transferQty & " to " & destinationID & "</p>"
		end function

	end class
%>