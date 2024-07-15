<p align="center">
  <img src="https://openclipart.org/image/800px/241044" />
</p>

# Description

This workbook shows how to integrate the Google Maps Directions API to calculate driving times between locations stored in an Excel spreadsheet.
The same concept can be applied to many APIs exposing an XML interface via HTTP.
The workbook contains three sheets: one with the location of some warehouses, one with the location of some customers, and one where the results will be written.
When the `ComputeClosestWarehouse` function is called, the script will determine the closest warehouse for each customer, based on the driving time between the locations, as returned by the Google Maps Directions API.
As this is a small tutorial, I did not add a graphical user interface to the workbook.

# VBA Code

I copy here all the VBA code used in the workbook. It consists of three functions:

* `CalculateDrivingTime` fetches the driving time between two addresses via the GMap API.
* `GetClosestWarehouseFor` finds the closest warehouse for a given customer (using `CalculateDrivingTime`).
* `ComputeClosestWarehouses` determines the closest warehouse for each customer (using `GetClosestWarehouseFor`).

```vb
Option Explicit

Sub ComputeClosestWarehouses()
    Dim customersTable As ListObject
    Set customersTable = Worksheets("Customers").ListObjects("CustomersTbl")
    
    Dim clsTable As ListObject
    Set clsTable = Worksheets("ClosestWarehouse").ListObjects("ClosestWarehouseTbl")
    
    Dim customerIdIndex As Integer
    Dim customerNameIndex As Integer
    Dim customerAddressIndex As Integer
    
    customerIdIndex = customersTable.ListColumns("Id").Index
    customerNameIndex = customersTable.ListColumns("Name").Index
    customerAddressIndex = customersTable.ListColumns("Address").Index
    
    Dim customerRow As ListRow

    For Each customerRow In customersTable.ListRows
        Dim findCustomer As Range
        Set findCustomer = clsTable.ListColumns("CustomerId").Range.Find(customerRow.Range(customerIdIndex))
        
        If findCustomer Is Nothing Then
            Dim closestWarehouseId As String
            Dim closestWarehouseName As String
            Dim closestWarehouseDrivingTime As Long
            
            GetClosestWarehouseFor customerRow.Range(customerAddressIndex), _
                                closestWarehouseId, _
                                closestWarehouseName, _
                                closestWarehouseDrivingTime
                                
            Dim newRow As ListRow
            Set newRow = clsTable.ListRows.Add
            
            With newRow
                .Range(1) = customerRow.Range(customerIdIndex)
                .Range(2) = customerRow.Range(customerNameIndex)
                .Range(3) = closestWarehouseId
                .Range(4) = closestWarehouseName
                .Range(5) = closestWarehouseDrivingTime
            End With
        End If
    Next customerRow
End Sub

Private Sub GetClosestWarehouseFor(ByVal customerAddress As String, _
                                ByRef closestWarehouseId As String, _
                                ByRef closestWarehouseName As String, _
                                ByRef closestWarehouseDrivingTime As Long)
    Dim bestWarehouseId As Variant
    Dim bestWarehouseName As Variant
    Dim bestDrivingTime As Variant
    bestDrivingTime = -1
        
    Dim warehousesTable As ListObject
    Set warehousesTable = Worksheets("Warehouses").ListObjects("WarehousesTbl")
    
    Dim warehouseIdIndex As Integer
    Dim warehouseNameIndex As Integer
    Dim warehouseAddressIndex As Integer
    
    warehouseIdIndex = warehousesTable.ListColumns("Id").Index
    warehouseNameIndex = warehousesTable.ListColumns("Name").Index
    warehouseAddressIndex = warehousesTable.ListColumns("Address").Index
    
    Dim warehouseRow As ListRow
    
    For Each warehouseRow In warehousesTable.ListRows
        Dim drivingTime As Long
        drivingTime = CalculateDrivingTime(warehouseRow.Range(warehouseAddressIndex), customerAddress)
        
        Debug.Print "Driving time between " & warehouseRow.Range(warehouseAddressIndex) & _
                    " and " & customerAddress & _
                    " is " & drivingTime
        
        If ((bestDrivingTime < 0) Or (bestDrivingTime > drivingTime)) Then
            bestWarehouseId = warehouseRow.Range(warehouseIdIndex)
            bestWarehouseName = warehouseRow.Range(warehouseNameIndex)
            bestDrivingTime = drivingTime
        End If
    Next warehouseRow
    
    closestWarehouseId = CStr(bestWarehouseId)
    closestWarehouseName = CStr(bestWarehouseName)
    closestWarehouseDrivingTime = CLng(bestDrivingTime)
End Sub

Private Function CalculateDrivingTime(ByVal warehouseAddress As String, ByVal customerAddress As String) As Long
    Dim url As String
    url = "https://maps.googleapis.com/maps/api/directions/xml" & _
        "?origin=" & WorksheetFunction.EncodeURL(warehouseAddress) & _
        "&destination=" & WorksheetFunction.EncodeURL(customerAddress)
        
    Dim request As New MSXML2.XMLHTTP60
    request.Open "GET", url, False
    request.send
    
    If (request.Status <> 200) Then
        Debug.Print "HTTP Status is not OK (200)"
        Debug.Print request.responseText
        Err.Raise 901, "CalculateDrivingTime", "The HTTP response status is not 200"
    End If
    
    Dim xmlDocument As MSXML2.DOMDocument60
    Set xmlDocument = request.responseXML
    
    Dim durationNode As MSXML2.IXMLDOMNode
    Set durationNode = xmlDocument.SelectSingleNode("/DirectionsResponse/route/leg/duration/value")
    
    If (durationNode Is Nothing) Then
        Debug.Print "Could not find the duration element in the XML document"
        Err.Raise 902, "CalculateDrivingTime", "The XML response did not contain a duration node"
    End If
    
    CalculateDrivingTime = CLng(durationNode.Text)
End Function
```

# License

This work is licensed under the GPL v3.0 license.
