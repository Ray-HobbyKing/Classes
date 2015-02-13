<!-- #include file = "../db.asp" -->
<!-- #include file = "../BackendSecurity.asp" -->
<!-- #include file = "backenduser.asp" -->
<!-- #include file = "shipping.asp" -->
<!-- #include file = "class_inTransit.asp" -->
<!-- #include file = "tableBuilder.asp" -->
<%
	dim title : title = "Class testing page"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
	<head>
		<title><%=title%></title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />
		<script src="http://code.jquery.com/jquery-1.9.1.js"></script>
		<script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
		<script src="../wps/js/jquery.tablesorter.min.js"></script>
		<script src="../wps/js/jquery.tablesorter.widgets.js"></script>
		<script src="../wps/js/jquery.tablesorter.pager.js"></script>
		<script type="text/javascript">
			$(document).ready(function() {
				$("#tableToSort").tablesorter({
					widgets: ['stickyHeaders'],
					headers: {
						'th': {
							sorter: false
						}
					}
				});
			});
		</script>
		<style type="text/css">
			table.collapse,
			table thead tr th.collapse,
			table tbody tr td.collapse {
				border-collapse: collapse;
			}
		</style>
	</head>
	<body>
		<h1><%=title%></h1>
		<form method="POST">
			<label for="idproduct">ID to Translate: </label>
			<input type="text" name="idproduct" id="idproduct" value="<%=request.form("idproduct")%>" />
			<input type="submit" name="Submit" value="Search" />
		</form>
		<br />
		<%
			'on error resume next
			''	set oUser = (new backenduser)(408)

			'	response.write oUser.idUser & "<br />"
			'	response.write oUser.getUsername & "<br />"

			'on error goto 0
			'set oShipment = (new shippingNumbers)(1)
			'oShipment.debug = 1
			'response.write "This is the old shipmentRef: " & oShipment.shipmentRef & "<br />"

			if request.form("Submit") = "Search" then
				set oProduct = (new barcodeTranslation)(request.form("idproduct"))
				oProduct.showAll()
			end if


			'set oTransfer = (new inTransit)("aaa")
			'oTransfer.destination("au")
			'oTransfer.transferQuantity(100)
			'oTransfer.confirmTransfer()

			set rs = Server.CreateObject("ADODB.Recordset")
				strsql =	"SELECT TOP 100 * FROM dbo.products WITH(NOLOCK) ORDER BY idproduct"
				rs.Open strsql, pDatabaseConnectionString

				set tb = (new tableBuilder)(rs)
					with tb
						.headerOrder   "idproduct,description,sku,cost"
						.relabelHeader "idproduct", "Barcode"
						.relabelHeader "sku", "SKU"
						.tableSortable = true
					end with

					response.write tb.render_as_table()
				set tb = nothing
				rs.close
			set rs = nothing
		%>
	</body>
</html>