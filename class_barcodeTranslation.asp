<%
	class barcodeTranslation
		public input_idproduct, conn, rs, debug

		public i_hkBarcode
		public i_arBarcode
		public i_auBarcode
		public i_cnBarcode
		public i_nlBarcode
		public i_ruBarcode
		public i_ukBarcode
		public i_usBarcode

		private sub Class_Initialize()
			set conn = Server.CreateObject("ADODB.Connection")
			conn.Open pDatabaseConnectionString

			i_hkBarcode = 0
			i_arBarcode = 0
			i_auBarcode = 0
			i_cnBarcode = 0
			i_nlBarcode = 0
			i_ruBarcode = 0
			i_ukBarcode = 0
			i_usBarcode = 0
			debug = 1
		end sub


		private sub Class_Terminate()
			conn.close
			set conn = nothing
		end sub


		public default function Init(seed)
			input_idproduct = seed
			if getBarcodes() = 0 then
				response.write "<p>Initialization failed</p>"
			end if
			set Init = Me
		end function

		public property get inputIdproduct()
			inputIdproduct = input_idproduct
		end property

		public property get hk()
			hk = i_hkBarcode
		end property

		public property get ar()
			ar = i_arBarcode
		end property

		public property get au()
			au = i_auBarcode
		end property

		public property get cn()
			cn = i_cnBarcode
		end property

		public property get nl()
			nl = i_nlBarcode
		end property

		public property get ru()
			ru = i_ruBarcode
		end property

		public property get uk()
			uk = i_ukBarcode
		end property

		public property get us()
			us = i_usBarcode
		end property


		private function getBarcodes()
			set rs = Server.CreateObject("ADODB.RecordSet")
				strsql =	"SELECT " & vbcrlf &_
							"	p.idproduct AS [HK ID], isnull(ar.idproduct, 0) AS [AR ID], isnull(au.idproduct, 0) AS [AU ID], " & vbcrlf &_
							"	isnull(cn.idproduct, 0) AS [CN ID], isnull(nl.idproduct, 0) AS [NL ID], isnull(ru.idproduct, 0) AS [RU ID], " & vbcrlf &_
							"	isnull(gb.idproduct, 0) AS [GB ID], isnull(us.idproduct, 0) AS [US ID] " & vbcrlf &_
							"FROM dbo.products AS p WITH(NOLOCK) " & vbcrlf &_
							"LEFT JOIN dbo.fproducts AS fp WITH(NOLOCK) ON fp.idproduct = p.idproduct " & vbcrlf &_
							"OUTER APPLY ( " & vbcrlf &_
							"	SELECT p1.idproduct, p1.sku " & vbcrlf &_
							"	FROM dbo.products AS p1 WITH(NOLOCK) " & vbcrlf &_
							"	INNER JOIN dbo.fproducts AS fp1 WITH(NOLOCK) ON fp1.idproduct = p1.idproduct AND fp1.countrycode = 'AR' " & vbcrlf &_
							"	WHERE p1.sku = p.sku " & vbcrlf &_
							") AS ar " & vbcrlf &_
							"OUTER APPLY ( " & vbcrlf &_
							"	SELECT p2.idproduct, p2.sku " & vbcrlf &_
							"	FROM dbo.products AS p2 WITH(NOLOCK) " & vbcrlf &_
							"	INNER JOIN dbo.fproducts AS fp2 WITH(NOLOCK) ON fp2.idproduct = p2.idproduct AND fp2.countrycode = 'AU' " & vbcrlf &_
							"	WHERE p2.sku = p.sku " & vbcrlf &_
							") AS au " & vbcrlf &_
							"OUTER APPLY ( " & vbcrlf &_
							"	SELECT p3.idproduct, p3.sku " & vbcrlf &_
							"	FROM dbo.products AS p3 WITH(NOLOCK) " & vbcrlf &_
							"	INNER JOIN dbo.fproducts AS fp3 WITH(NOLOCK) ON fp3.idproduct = p3.idproduct AND fp3.countrycode = 'CN' " & vbcrlf &_
							"	WHERE p3.sku = p.sku " & vbcrlf &_
							") AS cn " & vbcrlf &_
							"OUTER APPLY ( " & vbcrlf &_
							"	SELECT p4.idproduct, p.sku " & vbcrlf &_
							"	FROM dbo.products AS p4 WITH(NOLOCK) " & vbcrlf &_
							"	INNER JOIN dbo.fproducts AS fp4 WITH(NOLOCK) ON fp4.idproduct = p4.idproduct AND fp4.countrycode = 'NL' " & vbcrlf &_
							"	WHERE p4.sku = p.sku " & vbcrlf &_
							") AS nl " & vbcrlf &_
							"OUTER APPLY ( " & vbcrlf &_
							"	SELECT p5.idproduct, p5.sku " & vbcrlf &_
							"	FROM dbo.products AS p5 WITH(NOLOCK) " & vbcrlf &_
							"	INNER JOIN dbo.fproducts AS fp5 WITH(NOLOCK) ON fp5.idproduct = p5.idproduct AND fp5.countrycode = 'RU' " & vbcrlf &_
							"	WHERE p5.sku = p.sku " & vbcrlf &_
							") AS ru " & vbcrlf &_
							"OUTER APPLY ( " & vbcrlf &_
							"	SELECT p6.idproduct, p6.sku " & vbcrlf &_
							"	FROM dbo.products AS p6 WITH(NOLOCK) " & vbcrlf &_
							"	INNER JOIN dbo.fproducts AS fp6 WITH(NOLOCK) ON fp6.idproduct = p6.idproduct AND fp6.countrycode = 'GB' " & vbcrlf &_
							"	WHERE p6.sku = p.sku " & vbcrlf &_
							") AS gb " & vbcrlf &_
							"OUTER APPLY ( " & vbcrlf &_
							"	SELECT p7.idproduct, p7.sku " & vbcrlf &_
							"	FROM dbo.products AS p7 WITH(NOLOCK) " & vbcrlf &_
							"	INNER JOIN dbo.fproducts AS fp7 WITH(NOLOCK) ON fp7.idproduct = p7.idproduct AND fp7.countrycode = 'US' " & vbcrlf &_
							"	WHERE p7.sku = p.sku " & vbcrlf &_
							") AS us " & vbcrlf &_
							"WHERE fp.countrycode IS NULL " & vbcrlf &_
							"AND p.sku = ( " & vbcrlf &_
							"	SELECT sku FROM products WITH(NOLOCK) WHERE idproduct = " & inputIdproduct & " " & vbcrlf &_
							")"
				rs.open strsql, conn

				if not rs.eof then
					i_hkBarcode = rs("HK ID")
					i_arBarcode = rs("AR ID")
					i_auBarcode = rs("AU ID")
					i_cnBarcode = rs("CN ID")
					i_nlBarcode = rs("NL ID")
					i_ruBarcode = rs("RU ID")
					i_ukBarcode = rs("GB ID")
					i_usBarcode = rs("US ID")
					getBarcodes = 1
				else
					getBarcodes = 0
				end if
				rs.close
			set rs = nothing
		end function


		public function showAll()
			response.write "<pre>HK: " & hk & "</pre>"
			response.write "<pre>AU: " & au & "</pre>"
			response.write "<pre>AR: " & ar & "</pre>"
			response.write "<pre>CN: " & cn & "</pre>"
			response.write "<pre>NL: " & nl & "</pre>"
			response.write "<pre>RU: " & ru & "</pre>"
			response.write "<pre>UK: " & uk & "</pre>"
			response.write "<pre>US: " & us & "</pre>"
		end function

	end class
%>