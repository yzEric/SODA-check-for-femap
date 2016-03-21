'______________________________________________________________
'    Name:       SODA - Sum Of Deviation Angles
'    Author:     E. LE GAL
'    Version:    1.0
'    Date:       15/03/2016
'    Languages:  WinWrap
'    Purpose:    -
'______________________________________________________________
'
'
' The MIT License (MIT)
'
' Copyright (c) 2016 Eric LE GAL
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
'______________________________________________________________

Option Explicit    '-> All variables must be declared
Option Base 0      '-> First index of Arrays is 0



Private PI As Double
Private halfPI As Double
Private radToDeg_coeff As Double


Sub ini_const()
  PI = 4 * Atn(1)
  halfPI = 2 * Atn(1)
  radToDeg_coeff = 45/Atn(1) 
End Sub



Sub Main()

  Call ini_const

  Dim femapMod As femap.model , rc As Long
  Set femapMod = feFemap() 



'--- Aks User to select elements
  Dim elemSet As femap.Set
  Set elemSet = femapMod.feSet
  rc= elemSet.Select(FT_ELEM, True, "Select 2D elements")
      If rc= FE_CANCEL Then Exit Sub
      If AssertRC( femapMod, rc, "Unable to select elements" ) Then Exit Sub


'--- Start crono
  Dim startTime As Double
  startTime = Timer


'--- Keep only 2D shape elements
  Dim shapeSet As femap.Set, All2DElemsSet As femap.Set
  Set shapeSet = femapMod.feSet
  Set All2DElemsSet = femapMod.feSet

  rc= shapeSet.AddArray( 4, Array(FTO_TRIA3, FTO_TRIA6, FTO_QUAD4, FTO_QUAD8 ))
     If AssertRC( femapMod, rc, "Unable to add data to shapeSet" ) Then Exit Sub  


  rc= All2DElemsSet.AddSetRule( shapeSet.ID , FGD_Elem_byShape )
      If AssertRC( femapMod, rc, "Unable to create a set of element by using shapes" ) Then Exit Sub

  rc= elemSet.RemoveNotCommon( All2DElemsSet.ID )
      If AssertRC( femapMod, rc, "Unable to exclude non 2D elements" ) Then Exit Sub

  Set shapeSet = Nothing
  Set All2DElemsSet = Nothing




'--- Get Coordinates of nodes used by elements
  Call femapMod.feAppMessage(FCM_NORMAL, "1/2 : Get coordinate of nodes")

  Dim nodeSet As femap.Set, coordTable As Variant
  Set nodeSet = femapMod.feSet

  rc= nodeSet.AddSetRule( elemSet.ID , FGD_NODE_ONELEM )
      If AssertRC( femapMod, rc, "Uanble to get list of nodes used by elements" ) Then Exit Sub

  If Not GetNodesCoord_bigTable( femapMod, nodeSet, coordTable ) Then Exit Sub

  Set nodeSet = Nothing



'--- Calculation of distortions
  Call femapMod.feAppMessage(FCM_NORMAL, "2/2 : Calculation of criteria")

  Dim SODA() As Double, elemIDs() As Long
  If Not getDistortionSODA( femapMod, elemSet, coordTable, True, SODA, elemIDs ) Then Exit Sub

  Set elemSet = Nothing


'--- Create outputSet
  Const outputSet_Title= "Distortion SODA"
  Dim OutputSetID As Long
  OutputSetID = femapMod.Info_NextID( FT_OUT_CASE )

  If Not createOutputSet( femapMod, OutputSetID, outputSet_Title, "Sum of deviation angles", 0 ) Then Exit Sub


'--- Write results ---
  Const vectorID= 400000 '9000000
  Const vectorTitle= "Sum of deviation angle"

  Dim outVector As femap.Output, count As Long
  Set outVector = femapMod.feOutput
  count= UBound(elemIDs)+1 

  rc= outVector.InitScalarAtElem(OutputSetID, vectorID, vectorTitle, 10, False)
      If AssertRC( femapMod, rc, "Unable to initialize the output vector" ) Then Exit Sub

  rc= outVector.PutScalarAtElem(count, elemIDs, SODA)
      If AssertRC( femapMod, rc, "Unable to put results in the output vector" ) Then Exit Sub

  rc= outVector.Put(0)
      If AssertRC( femapMod, rc, "Unable to put the output vector in the model" ) Then Exit Sub

  Call femapMod.feAppMessage( FCM_NORMAL, "Output Set " + CStr(OutputSetID) + ": " + CStr(vectorID) + " -> " + vectorTitle)


'--- This is The End
  femapMod.feAppMessage( FCM_NORMAL ,"--- End of Task " +Format((Timer-startTime) , "0.0s"))
  Set femapMod = Nothing
End Sub


'--------------------------------------------------------------------------


Function GetNodesCoord_bigTable( femapMod As femap.model, nodeSet As femap.Set, outTable As Variant )As Boolean
'
' - Description
'     This function create a table of node coordinates
'
' - Input:
'     femapMod -> femap model to use
'     nodeSet  -> femap set of node to use
'
' - output:
'     outTable[0..2, 0..maxNodeID] -> A 2D table of node coordinates
'
' - return code
'    True  -> OK, no problem
'    False -> Something goes wrong

 If nodeSet.Count= 0 Then
   GetNodesCoord_bigTable= True : Exit Function
 End If

 Dim aNode As femap.Node, nodeCount As Long, xyz As Variant, IDs As Variant, maxID As Long, rc As Long
 Set aNode = femapMod.feNode

 rc= aNode.GetCoordArray( nodeSet.ID, nodeCount, IDs, xyz )
       If AssertRC( femapMod, rc, "Unable to obtain coordinates of nodes") Then Exit Function

 maxID= nodeSet.Last
 ReDim outTable( 2 , maxID )

 Dim i As Long, nodes_Last As Long, nodeID As Long, elemIndex As Long
 nodes_Last =  nodeCount-1

 For i=0 To nodes_Last
    nodeID = IDs( i )
    outTable( 0, nodeID )= xyz(elemIndex)
    outTable( 1, nodeID )= xyz(elemIndex+1)
    outTable( 2, nodeID )= xyz(elemIndex+2)
    elemIndex=elemIndex+3
 Next i

 GetNodesCoord_bigTable=True
End Function


'--------------------------------------------------------------------------


Function GetElemDistortionSODA_Quad( _
 ByVal x0 As Double , ByVal y0 As Double , ByVal z0 As Double , _
 byval x1 as double , byval y1 as double , byval z1 as double , _
 ByVal x2 As Double , ByVal y2 As Double , ByVal z2 As Double , _
 byval x3 as double , byval y3 as double , byval z3 as double ) as double
'
' - Description
'     This function calculate sum of deviation angles for a quad element
'
' - Input:
'     x0, y0, z0 -> coordinates of node 0
'     x1, y1, z1 -> coordinates of node 1
'     ...
'
' - output:
'     The result value
'
' - Remarks/Usage:
'     Value of the 4th angle is based on the values of the 3 others, 
'     -> imprecise with very high warp 

 Dim angle_301 As Double, angle_012 As Double, angle_123 As Double, angle_230 As Double
 angle_301 = AngleOf2Vectors_deg( x3-x0 , y3-y0 , z3-z0 , x1-x0 , y1-y0 , z1-z0 )
 angle_012 = AngleOf2Vectors_deg( x0-x1 , y0-y1 , z0-z1 , x2-x1 , y2-y1 , z2-z1 )
 angle_123 = AngleOf2Vectors_deg( x1-x2 , y1-y2 , z1-z2 , x3-x2 , y3-y2 , z3-z2 )
 angle_230 = 360.0 - angle_301 - angle_012 - angle_123


 GetElemDistortionSODA_Quad = Abs(angle_301-90.0) +Abs(angle_012-90.0)+Abs(angle_123-90.0)+Abs(angle_230-90.0)
End Function



Function GetElemDistortionSODA_Tria( _
 ByVal x0 As Double , ByVal y0 As Double , ByVal z0 As Double , _
 ByVal x1 As Double , ByVal y1 As Double , ByVal z1 As Double , _
 ByVal x2 As Double , ByVal y2 As Double , ByVal z2 As Double ) As Double
'
' - Description
'     This function calculate sum of deviation angles for a tria element
'
' - Input:
'     x0, y0, z0 -> coordinates of node 0
'     x1, y1, z1 -> coordinates of node 1
'     x2, y2, z2 -> coordinates of node 2
'
' - output:
'     The result value
'
' - Remarks/Usage:
'

 Dim angle_201 As Double, angle_012 As Double, angle_123 As Double
 angle_201 = AngleOf2Vectors_deg( x2-x0 , y2-y0 , z2-z0 , x1-x0 , y1-y0 , z1-z0 )
 angle_012 = AngleOf2Vectors_deg( x0-x1 , y0-y1 , z0-z1 , x2-x1 , y2-y1 , z2-z1 )
 angle_123 = 180.0 - angle_201 - angle_012

 GetElemDistortionSODA_Tria = Abs(angle_201-60.0) +Abs(angle_012-60.0)+Abs(angle_123-60.0)
End Function


'--------------------------------------------------------------------------


Function getDistortionSODA(femapMod As femap.model, elemSet As femap.Set, coordTable As Variant, showStatusBar As Boolean, outSODA() As Double, outIDs() As Long) As Boolean
'
' - Description
'     This function calculate sum of deviation angles of elements quad or tria
'
' - Input:
'     femapMod -> Femap model to use
'     ElemSet  -> Set of element to use
'     coordTable(0..2, 0..maxNodeID) -> coordinates of nodes
'     showStatusBar
'
' - output:
'     outSODA(0..elemCount-1)
'     outIDs(0..elemCount-1)  
'
' - return code
'    True  -> OK, no problem
'    False -> Something goes wrong
'
' - Remarks/Usage:
'     for non 2D elements SODA value is 0

  If elemSet.Count=0 Then
    getDistortionSODA=True : Exit Function
  End If


'--- Get data of elements : nodes
   Dim anElem As femap.Elem, rc As Long
   Set anElem = femapMod.feElem

   Dim numElem As Long, entID As Variant, propID As Variant, elemTYPE As Variant, topology As Variant, layerID As Variant
   Dim color As Variant, formulation As Variant, orient As Variant, offset As Variant, release As Variant
   Dim orientSET As Variant, orientID As Variant, ElemNodes As Variant, connectTYPE As Variant, connectSEG As Variant

   rc= anElem.GetAllArray( elemSet.ID, numElem, entID, propID, elemTYPE, topology, layerID, color, formulation, orient, offset, release, orientSET, orientID, ElemNodes, connectTYPE, connectSEG )
       If AssertRC( femapMod, rc, "Unable to obtain data of elements" ) Then Exit Function


   Erase propID, elemTYPE, layerID, color, formulation, orient, offset, release, orientSET, orientID, connectTYPE, connectSEG
   'keep: entID, topology, ElemNodes 



'--- Compute SODA

  Dim lastElem As Long, i_Elem As Long
  lastElem = numElem-1
  ReDim outSODA(lastElem) As Double, outIDs(lastElem) As Long

  Dim elemIndex As Long, n0 As Long, n1 As Long, n2 As Long, n3 As Long
  'elemIndex= 0


  Dim statusInc as long , statusUpdateStep as Long
  If showStatusBar Then
     statusUpdateStep= CLng(lastElem/25 + 0.49999999)
     Call FemapMod.feAppStatusShow( True, lastElem )
  End If
  
  
  For i_Elem=0 To lastElem
     outIDs(i_Elem)= entID(i_Elem)

     Select Case topology(i_Elem)

     Case FTO_QUAD4, FTO_QUAD8
          n0=ElemNodes(elemIndex)
          n1=ElemNodes(elemIndex+1)
          n2=ElemNodes(elemIndex+2)
          n3=ElemNodes(elemIndex+3)
          outSODA(i_Elem)=GetElemDistortionSODA_Quad( _
              coordTable(0, n0 ) , coordTable(1, n0 ) , coordTable(2, n0 ) , _
              coordTable(0, n1 ) , coordTable(1, n1 ) , coordTable(2, n1 ) , _
              coordTable(0, n2 ) , coordTable(1, n2 ) , coordTable(2, n2 ) , _
              coordTable(0, n3 ) , coordTable(1, n3 ) , coordTable(2, n3 ) )
        
     Case FTO_TRIA3, FTO_TRIA6
          n0=ElemNodes(elemIndex)
          n1=ElemNodes(elemIndex+1)
          n2=ElemNodes(elemIndex+2)
          outSODA(i_Elem)=GetElemDistortionSODA_Tria( _
              coordTable(0, n0 ) , coordTable(1, n0 ) , coordTable(2, n0 ) , _
              coordTable(0, n1 ) , coordTable(1, n1 ) , coordTable(2, n1 ) , _
              coordTable(0, n2 ) , coordTable(1, n2 ) , coordTable(2, n2 ) )
        
     Case Else
          'outSODA(i_Elem)=0
          
     End Select


     elemIndex=elemIndex +20
     
     If showStatusBar Then '-> Increase StatusBar
        statusInc=statusInc+1
        If statusInc=statusUpdateStep Then
           femapMod.feAppStatusUpdate( i_Elem )  : statusInc=1
        End If
     End If
        
    Next i_Elem
    Call FemapMod.feAppStatusShow(False, 0) '-> Hide StatusBar
     
   getDistortionSODA= True
End Function


'--------------------------------------------------------------------------


Function acos(ByVal x As Double) As Double
'
' - Description
'     This function returns the arc cos value of an angle
'
' - Input:
'     x -> value between -1 and 1
'
' - output:
'     The arc cos value of an angle
'
' - Remarks/Usage:
'   acos = Atn(-x / Sqr(1 - x^2)) + 2 * Atn(1)

 acos = Atn(-x / Sqr(1 - x^2)) + halfPI
 
End Function


'--------------------------------------------------------------------------


Function AngleOf2Vectors_deg( _
          byval Xa As Double, byval Ya As Double, byval Za As Double, _
          byval Xb As Double, byval Yb As Double, byval Zb As Double )As Double
'
' - Description
'     This function returns the angle of 2 vectors in degree
'
' - Input:
'     Xa, Ya, Za -> components of vector A
'     Xb, Yb, Zb -> components of vector B
'
' - output:
'     The value of the angle in degree
'
' - Remarks/Usage:
'   Input vectors A & B doesn't need to be normalized
'   cos angle = (Xa.Xb+Ya.Yb+Za.Zb) / sqrt((Xa²+Ya²+Za²)(Xb²+Yb²+Zb² ))

  On Error GoTo errorInOp ' It is faster to catch error divBy0 Than to prevent it 

  AngleOf2Vectors_deg = radToDeg_coeff * acos(  (Xa*Xb+Ya*Yb+Za*Zb)/Sqr( (Xa^2+Ya^2+Za^2)*(Xb^2+Yb^2+Zb^2) )  )
  Exit Function

  errorInOp:
  If (Err.Number =10061 ) Then
     Err.Clear
     AngleOf2Vectors_deg = 0
  Else
     Err.raise( Err.Number )   
  End If

End Function


'--------------------------------------------------------------------------


Function createOutputSet( femapMod As femap.model ,ID As Long, Title As String, notes As String, value As Double ) As Boolean
'
' - Description
'     This function create an outputSet in a model
'
' - Input:
'     femapMod -> Femap model to use
'     ID       -> ID of the outputset to create
'     Title    -> Title of the outputset to create
'     notes    -> notes of the outputset to create
'     value    -> value of the outputset to create
'
' - output:
'
' - return code
'    True  -> OK, no problem
'    False -> Something goes wrong
'
' - Remarks/Usage:
'
  Dim anOutputSet As femap.OutputSet, rc As Long
  Set anOutputSet = femapMod.feOutputSet

  anOutputSet.title=  Title
  anOutputSet.notes = notes
  anOutputSet.Value = value
  anOutputSet.program =FAP_FEMAP_GEN
  anOutputSet.analysis = FAT_UNKNOWN

  rc= anOutputSet.Put(ID)
     If AssertRC( femapMod, rc, "Unable to create the OutputSet" ) Then Exit Function

  createOutputSet = True
End Function


'--------------------------------------------------------------------------


Function AssertRC( femapMod As femap.model, rc As Long, msg As String )As Boolean
'
' - Description
'     This function display informations about return code when its value is different from FE_OK
'     If rc= FE_OK function return false
'
' - Input:
'     femapMod -> Femap model for message
'     rc       -> the return code to check
'     msg      -> message to show if return code is not FE_OK
'
' - output:
'
' - return code
'    True  -> OK, rc is equal to FE_OK
'    False -> rc is not equal to FE_OK
'
' - Remarks/Usage:
'    rc= femap.functionWithARetunCode( fewParameters )
'        If AssertRC( femapLink, rc, "The function 'functionWithARetun' procude an Error" ) Then
'          deal with error or maybe just exit
'        End If
'
  If rc=FE_OK Then Exit Function

  Dim info As String
  Select Case rc
  Case FE_TOO_SMALL
    info = "Too small"
  Case FE_FAIL
    info = "Fail"
  Case FE_BAD_TYPE
    info = "Bad type"
  Case FE_CANCEL
    info = "Cancel"
  Case FE_BAD_DATA
    info = "Bad data"
  Case FE_INVALID
    info = "Invalid"
  Case FE_NO_MEMORY
    info = "No memory"
  Case FE_NOT_EXIST
    info = "Not Exist"
  Case FE_NEGATIVE_MASS_VOLUME
    info = "Negative mass volume"
  Case FE_SECURITY
    info = "Security"
  Case FE_NO_FILENAME
    info = "No file name"
  Case FE_NOT_AVAILABLE
    info = "Not available"
  Case Else
    info = "Unkonw return code: "+CStr(rc)
  End Select

  info = "Return code is '" +info+ "'"

  On Error Resume Next

  Call femapMod.feAppMessage(FCM_ERROR, msg )
  Call femapMod.feAppMessage(FCM_ERROR, info )

  Call MsgBox( msg + vbCrLf + info )
  AssertRC = True
End Function
