Attribute VB_Name = "ConvertXML"

' GenerateXMLDOM
' @brief Generate an MS XML Object (without any format tags) based on the data inside selected region on the excel sheet
'
' @parameters rngData : The selected region on excel sheet, with the first row as field name, and data rows below
'                       For the field name, use the node delimiter "/" to build the hierarchy of data
'                       e.g. /data/field1 is equvalent to <data><field1>....</field1><data>
'
'             rootNodeName : The xml document root node tag name
'
' @return an MS XML Object
'
' @author  Raymond Pang
' @version 0.8

Function GenerateXMLDOM(rngData As Range, rootNodeName As String)

    Const NODE_DELIMITER As String = "/"   ' the default node delimiter
    
    Dim intColCount As Integer
    Dim intRowCount As Integer
    Dim intColCounter As Integer
    Dim intRowCounter As Integer
    
       
    Dim rngCell As Range
  
    ' Create the XML DOM object
    Set ObjXMLDoc = CreateObject("Microsoft.XMLDOM")
    ObjXMLDoc.async = False
                   
    
    ' NODE_PROCESSING_INSTRUCTION(7) --- reference http://www.devguru.com/Technologies/xmldom/quickref/obj_node.html
    Set Heading = ObjXMLDoc.createNode(7, "xml", "")
    ObjXMLDoc.appendChild (Heading)
    
    ' Set the root node
    Set top_node = ObjXMLDoc.createNode(1, rootNodeName, "")
    ObjXMLDoc.appendChild (top_node)

    
        
    Dim Nodes() As String       'Array storing the current splited node names
    Dim NodeStack() As String   'Array storing the last node names
    
    Dim new_nodes()
    ReDim NodeStack(0)
    ReDim new_nodes(0)
        
    
     With rngData  ' The selected region on the Excel Sheet passed in
        
        '   Discover dimensions of the data we  will be dealing with...
        intColCount = .Columns.Count
        intRowCount = .Rows.Count
        
        Dim strColNames() As String     ' The Array of column names
        ReDim strColNames(intColCount)
        
        
        ' First Row is the Field/Tag names
        ' Extract all the field names into array "strColNames"
        If intRowCount >= 1 Then
            '   Loop accross columns... and put names in array
            For intColCounter = 1 To intColCount

                '   Mark the cell under current scrutiny by setting
                '   an object variable...
                Set rngCell = .Cells(1, intColCounter)
                
                '  not support merged cells .. so quit
                If Not rngCell.MergeArea.Address = _
                                            rngCell.Address Then
                      MsgBox ("!! Cell Merged ... Invalid format")
                      Exit Function
                End If
                
                strColNames(intColCounter) = rngCell.Text
                                 
            Next
        End If
        
        '   Loop down the table's rows
        For intRowCounter = 2 To intRowCount
           
            ReDim new_nodes(0)
            ReDim NodeStack(0)
            '   Loop accross columns...
            For intColCounter = 1 To intColCount
                
                '   Mark the cell under current scrutiny by setting
                '   an object variable...
                Set rngCell = .Cells(intRowCounter, intColCounter)
                
                
                '   Is the cell merged?..
                If Not rngCell.MergeArea.Address = _
                                            rngCell.Address Then
                
                      MsgBox ("!! Cell Merged ... Invalid format")
                      Exit Function
                      
                                       
                End If
                          
                ' divide the field name by the delimiter to get appropriate node names
                Nodes = Split(strColNames(intColCounter), NODE_DELIMITER)
                    
                If UBound(Nodes) = 0 Then
                  ReDim Nodes(1)
                  Nodes(1) = strColNames(intColCounter)
                End If
                
                 ' don't count it when no content
                 If Trim(rngCell.Text) <> "" Then
                                    
                    Dim i As Integer
                    MatchAll = True
                    For i = 1 To UBound(Nodes)

                        If i <= UBound(NodeStack) Then
                            
                            If Trim(Nodes(i)) <> Trim(NodeStack(i)) Then
                                'not match
                                MatchAll = False
                                Exit For
                            
                            End If
                        Else
                          MatchAll = False
                          Exit For
                        End If
                       
                                                                          
                      Next
                      
                      ' match all means in same level as previous, so it needs to output for the last node
                      If MatchAll Then
                          i = i - 1
                      End If
                      
                          
                      If UBound(new_nodes) < UBound(Nodes) Then
                             ' enlong the array
                             ReDim Preserve new_nodes(UBound(Nodes))
                              
                              
                      End If
                          
                      For t = i To UBound(Nodes)
                          ' create uncommon nodes with the previous one
                          Set new_nodes(t) = ObjXMLDoc.createNode(1, Nodes(t), "")
                                                       
                      Next
                      
                      
                      For t = i - 1 To UBound(Nodes) - 1
                      
                          If t >= 1 Then
                              ' connect the nodes based on the hierarchy
                              new_nodes(t).appendChild (new_nodes(t + 1))
                          End If
                      
                      Next
                      Set Textcont = ObjXMLDoc.createTextNode(Trim(rngCell.Text))
                      new_nodes(UBound(Nodes)).appendChild (Textcont)
                       
                      If i = 1 Then
                          top_node.appendChild (new_nodes(1))
                      End If
                  
                  
                      NodeStack = Nodes
                         
                         
                                      
                  End If
               
            Next ' finished a column
        Next
    End With
    

    '   Return the XMLDOM
     Set GenerateXMLDOM = ObjXMLDoc
     
End Function



' fGenerateXML
' @brief: Generate a 'clean' XML (ie. no unwanted formatting tags)
'         from an Excel range.
'
' @parameters rngData : The selected region on excel sheet, with the first row as field name, and data rows below
'                       For the field name, apart from normal
'             rootNodeName : The xml document root node tag name
'
' @return String with the content of XML preparing to write out to file
'
' @author  Raymond Pang
' @version 0.8

Function fGenerateXML(rngData As Range, rootNodeName As String) As String

'===============================================================
'   XML Tags
    '   Table
    
    Const HEADER                As String = "<?xml version=""1.0""?>"
    Dim TAG_BEGIN  As String
    Dim TAG_END  As String
    Const NODE_DELIMITER        As String = "/"
    
        
'===============================================================

    Dim intColCount As Integer
    Dim intRowCount As Integer
    Dim intColCounter As Integer
    Dim intRowCounter As Integer
    
   
    Dim rngCell As Range
 
    
    Dim strXML As String
    
       
    
    '   Initial table tag...
    
    
   TAG_BEGIN = vbCrLf & "<" & rootNodeName & ">"
   TAG_END = vbCrLf & "</" & rootNodeName & ">"
   
    strXML = HEADER
    strXML = strXML & TAG_BEGIN
                    
    With rngData
        
        '   Discover dimensions of the data we
        '   will be dealing with...
        intColCount = .Columns.Count
        
        intRowCount = .Rows.Count
        
        Dim strColNames() As String
        
        ReDim strColNames(intColCount)
        
        
        ' First Row is the Field/Tag names
        If intRowCount >= 1 Then
        
            '   Loop accross columns...
            For intColCounter = 1 To intColCount
                
                '   Mark the cell under current scrutiny by setting
                '   an object variable...
                Set rngCell = .Cells(1, intColCounter)
                
              
                
                '   Is the cell merged?..
                If Not rngCell.MergeArea.Address = _
                                            rngCell.Address Then
               
                      MsgBox ("!! Cell Merged ... Invalid format")
                      Exit Function
                      
                                       
                End If
            
                 strColNames(intColCounter) = rngCell.Text
                
            Next
        
        End If
        
        
        Dim Nodes() As String
        Dim NodeStack() As String
      
      
        '   Loop down the table's rows
        For intRowCounter = 2 To intRowCount
           
            
            strXML = strXML & vbCrLf & TABLE_ROW
            ReDim NodeStack(0)
            '   Loop accross columns...
            For intColCounter = 1 To intColCount
                
                '   Mark the cell under current scrutiny by setting
                '   an object variable...
                Set rngCell = .Cells(intRowCounter, intColCounter)
                
                               
                '   Is the cell merged?..
                If Not rngCell.MergeArea.Address = _
                                            rngCell.Address Then
                
                      MsgBox ("!! Cell Merged ... Invalid format")
                      Exit Function
                      
                End If
               
                If Left(strColNames(intColCounter), 1) = NODE_DELIMITER Then
                      
                      Nodes = Split(strColNames(intColCounter), NODE_DELIMITER)
                          ' check whether we are starting a new node or not
                          Dim i As Integer
                         
                          Dim MatchAll As Boolean
                          MatchAll = True
                         
                          
                          For i = 1 To UBound(Nodes)
    
                              If i <= UBound(NodeStack) Then
                                  
                                  If Trim(Nodes(i)) <> Trim(NodeStack(i)) Then
                                      'not match
                                      'MsgBox (Nodes(i) & "," & NodeStack(i))
                                      MatchAll = False
                                      Exit For
                                  
                                  End If
                              Else
                                MatchAll = False
                                Exit For
                              End If
                              
                              
                                                                                
                          Next
                          
                          ' add close tags to those not used afterwards
                          
              
                         ' don't count it when no content
                         If Trim(rngCell.Text) <> "" Then
                            
                            If MatchAll Then
                              strXML = strXML & "</" & NodeStack(UBound(NodeStack)) & ">" & vbCrLf
                            Else
                              For t = UBound(NodeStack) To i Step -1
                                strXML = strXML & "</" & NodeStack(t) & ">" & vbCrLf
                              Next
                            End If
                            
                            If i < UBound(Nodes) Then
                                For t = i To UBound(Nodes)
                                    ' add to the xml
                                    strXML = strXML & "<" & Nodes(t) & ">"
                                    If t = UBound(Nodes) Then
                                                                           
                                            strXML = strXML & Trim(rngCell.Text)
                                        
                                    End If
                                    
                                Next
                              Else
                                  t = UBound(Nodes)
                                  ' add to the xml
                                  strXML = strXML & "<" & Nodes(t) & ">"
                                  strXML = strXML & Trim(rngCell.Text)

                              End If
                           
                              NodeStack = Nodes
                           
                          Else
                          
                            ' since its a blank field, so no need to handle if field name repeated
                            If Not MatchAll Then
                              For t = UBound(NodeStack) To i Step -1
                                strXML = strXML & "</" & Trim(NodeStack(t)) & ">" & vbCrLf
                              Next
                            End If
                            
                            ReDim Preserve NodeStack(i - 1)
                          End If
                            
                                              
                          ' the last column
                          If intColCounter = intColCount Then
                           ' add close tags to those not used afterwards
                              If UBound(NodeStack) <> 0 Then
                               For t = UBound(NodeStack) To 1 Step -1
                          
                              strXML = strXML & "</" & Trim(NodeStack(t)) & ">" & vbCrLf
                              
                              Next
                              End If
                          End If
                   
                 Else
                      ' add close tags to those not used afterwards
                      If UBound(NodeStack) <> 0 Then
                          For t = UBound(NodeStack) To 1 Step -1
                          
                           strXML = strXML & "</" & Trim(NodeStack(t)) & ">" & vbCrLf
                              
                          Next
                      End If
                      ReDim NodeStack(0)
      
                        ' skip if no content
                      If Trim(rngCell.Text) <> "" Then
                        strXML = strXML & "<" & Trim(strColNames(intColCounter)) & ">" & Trim(rngCell.Text) & "</" & Trim(strColNames(intColCounter)) & ">" & vbCrLf
                      End If
                      
                  End If
                
                    
                
                
            Next
           
        Next
    End With
    
    strXML = strXML & TAG_END
    
    '   Return the HTML string...
    fGenerateXML = strXML
    
End Function



' Function for writing plain string out a file

Sub sWriteFile(strXML As String, strFullFileName As String)

    Dim intFileNum As String
    
    intFileNum = FreeFile
    Open strFullFileName For Output As #intFileNum
    Print #intFileNum, strXML
    Close #intFileNum
    
    
End Sub

' To automatically select the "REAL"/non empty continuous regions (rows and columns)

Sub FindUsedRange()
    Dim LastRow As Long
    Dim FirstRow As Long
    Dim LastCol As Integer
    Dim FirstCol As Integer

    ' Find the FIRST real row
    FirstRow = ActiveSheet.Cells.Find(What:="*", _
      SearchDirection:=xlNext, _
      SearchOrder:=xlByRows).Row
      
    ' Find the FIRST real column
    FirstCol = ActiveSheet.Cells.Find(What:="*", _
      SearchDirection:=xlNext, _
      SearchOrder:=xlByColumns).Column
    
    ' Find the LAST real row
    LastRow = ActiveSheet.Cells.Find(What:="*", _
      SearchDirection:=xlPrevious, _
      SearchOrder:=xlByRows).Row

    ' Find the LAST real column
    LastCol = ActiveSheet.Cells.Find(What:="*", _
      SearchDirection:=xlPrevious, _
      SearchOrder:=xlByColumns).Column
        
'Select the ACTUAL Used Range as identified by the
'variables identified above
    'MsgBox (FirstRow & "," & LastRow & "," & FirstCol & "," & LastCol)
    Dim topCel As Range
    Dim bottomCel As Range
   
    Set topCel = Cells(FirstRow, FirstCol)
    Set bottomCel = Cells(LastRow, LastCol)
    
   ActiveSheet.Range(topCel, bottomCel).Select
End Sub


