'______________________________________________________________
'    Name:       SODA - Sum Of Deviation Angles
'    Author:     E. LE GAL
'    Version:    1.0
'    Date:       20/03/2016
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



Sub Main( )
  Dim femapMod As femap.model, rc As Long
  Set femapMod = feFemap()   ' A partir de la v10

'--- Aks User to select elements
  Dim elemSet As femap.Set, UserSelectCount As Long
  Set elemSet = femapMod.feSet
  rc= elemSet.Select(FT_ELEM, True, "Select 2D elements")
      If rc= FE_CANCEL Then Exit Sub
      If AssertRC( femapMod, rc, "Unable to select elements" ) Then Exit Sub

  UserSelectCount = elemSet.Count
 
'--- Start crono
  Dim startTime As Double
  startTime = Timer


'--- Keep only 2D shape elements
  Dim shapeSet As femap.Set, All2DElemsSet As femap.Set
  Set shapeSet = femapMod.feSet
  Set All2DElemsSet = femapMod.feSet

  rc= shapeSet.AddArray( 4, Array(FTO_TRIA3, FTO_TRIA6, FTO_QUAD4, FTO_QUAD8 ) )
     If AssertRC( femapMod, rc, "Unable to add data to shapeSet" ) Then Exit Sub 


  rc= All2DElemsSet.AddSetRule( shapeSet.ID , FGD_Elem_byShape )
      If AssertRC( femapMod, rc, "Unable to create a set of element by using shapes" ) Then Exit Sub

  rc= elemSet.RemoveNotCommon( All2DElemsSet.ID )
      If AssertRC( femapMod, rc, "Unable to exclude non 2D elements" ) Then Exit Sub

  Set shapeSet = Nothing
  Set All2DElemsSet = Nothing

  If UserSelectCount <> elemSet.Count Then femapMod.feAppMessage(FCM_HIGHLIGHT, CStr(elemSet.Count) + " elements removed from selection")

  If elemSet.Count = 0 Then
      femapMod.feAppMessage(femap.zMessageColor.FCM_HIGHLIGHT, "No selected element can be used") : Exit Sub
  End If

'--- Calculation of distortions
  Dim SODA() As Double, elemIDs() As Long
  If Not GetElemDistortionSODA( femapMod, elemSet, True, elemIDs, SODA ) Then Exit Sub

'--- Create outputSet
  Const outputSet_Title= "Distortion SODA"
  Dim OutputSetID As Long
  OutputSetID = femapMod.Info_NextID( FT_OUT_CASE )

  If Not createOutputSet( femapMod, OutputSetID, "Distortion SODA", "Sum of deviation angles", 0 ) Then Exit Sub


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


Function GetElemDistortionSODA( femapMod As femap.model, elemSet As femap.Set, showStatusBar As Boolean, ByRef elems_IDs As Variant, ByRef SODA() As Double ) As Long
'LastElem dernier Index d'élément
'ElemNodesIndex( e , n )  >> renvoie l'index du noeud "index n" de l'élément "index e"
'coordTable( u, e ) >> renvoie la ccordonnée u de l'élément "index e" ( u: 0 -> X , 1 -> Y , 2 -> Z )
'showStatusBar >> afficher la barre de progression de Femap

   Dim rc As Long, elemsCount As Long, entIDs As Variant
   rc= elemSet.GetArray(elemsCount ,entIDs )
       If AssertRC( femapMod, rc, "Unable to get list of elements from set" ) Then Exit Function
       
   If elemsCount= 0 Then
     Erase elems_IDs, SODA: GetElemDistortionSODA = True
     Exit Function
   End If

   Dim LastElem As Long, i_Elem As Long, tmpAngle As Double
   LastElem= elemsCount -1
   ReDim SODA(LastElem),elems_IDs(LastElem)

   Dim Elem As Object , elemID As Long , vNodes_IDs As Variant
   Set Elem = femapMod.feElem


   If showStatusBar Then Call StatusBar( LastElem+1, femapMod ) '-> Set max of StatusBar
   For i_Elem=0 To LastElem
   
        elemID =entIDs( i_Elem )
        elems_IDs(i_Elem)=elemID
        rc= Elem.Get(elemID)
            If rc<>FE_OK Then
               Call AssertRC( femapMod, rc, "Unable to get data of element "+CStr(elemID)) : Exit Function
            End If
        
        vNodes_IDs = Elem.vnode
   
        Select Case Elem.topology
        Case FTO_QUAD4, FTO_QUAD8
          If Not GetElemDistortionSODA_Quad( femapMod, vNodes_IDs, tmpAngle ) Then
            Call StatusBar( -1 ) '-> Hide StatusBar
            Exit Function
          End If
        Case FTO_TRIA3, FTO_TRIA6
          If Not GetElemDistortionSODA_Tria( femapMod, vNodes_IDs, tmpAngle ) Then
          	Call StatusBar( -1 ) '-> Hide StatusBar
            Exit Function
          End If
        Case Else ' bad element shape
            tmpAngle=0
        End Select

        SODA( i_Elem )=tmpAngle
        If showStatusBar Then Call StatusBar( 0 ) '-> Increase StatusBar
   Next i_Elem
   If showStatusBar Then Call StatusBar( -1 ) '-> Hide StatusBar

   GetElemDistortionSODA=True
End Function


'--------------------------------------------------------------------------


Function GetElemDistortionSODA_Quad( femapMod As femap.model, Nodes As Variant, outAngle As Double ) As Boolean
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
'
 Dim angle_301 As Double, angle_012 As Double, angle_123 As Double, angle_230 As Double
 Dim vecBase As Variant, vecNormal As Variant, rc As Long
 
 'angle_301
 rc= femapMod.feMeasureAngleBetweenNodes( Nodes(0), Nodes(3), Nodes(1), 0, 0, 0, vecBase, vecNormal, angle_301 )
 
 'angle_012
 If rc=FE_OK Then rc= femapMod.feMeasureAngleBetweenNodes( Nodes(1), Nodes(0), Nodes(2), 0, 0, 0, vecBase, vecNormal, angle_012 )

 'angle_123
 If rc=FE_OK Then rc= femapMod.feMeasureAngleBetweenNodes( Nodes(2), Nodes(1), Nodes(3), 0, 0, 0, vecBase, vecNormal, angle_123 )
 
 If rc<>FE_OK Then
    femapMod.feAppMessage( FCM_ERROR, "Unable to evaluate angles of corners angle of tria: "+CStr(Nodes(0))+","+CStr(Nodes(1))+","+CStr(Nodes(2))+","+CStr(Nodes(3)))
    Exit Function
 End If
 
 angle_230 = 360 - angle_301 - angle_012 - angle_123
 
 outAngle = Abs(angle_301-90.0) +Abs(angle_012-90.0) +Abs(angle_123-90.0) +Abs(angle_230-90.0)
 GetElemDistortionSODA_Quad = True
End Function


'--------------------------------------------------------------------------


Function GetElemDistortionSODA_Tria( femapMod As femap.model, Nodes As Variant, ByRef outAngle As Double ) As Double
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
 Dim vecBase As Variant, vecNormal As Variant, rc As Long
 
 'angle_201
 rc= femapMod.feMeasureAngleBetweenNodes( Nodes(0), Nodes(2), Nodes(1), 0, 0, 0, vecBase, vecNormal, angle_201 )
 
 'angle_012
 If rc= FE_OK Then rc= femapMod.feMeasureAngleBetweenNodes( Nodes(1), Nodes(0), Nodes(2), 0, 0, 0, vecBase, vecNormal, angle_012 )

 If rc<>FE_OK Then
    femapMod.feAppMessage( FCM_ERROR, "Unable to evaluate angles of corners angle of tria: "+CStr(Nodes(0))+","+CStr(Nodes(1))+","+CStr(Nodes(2)))
    Exit Function
 End If

 angle_123 = 180 - angle_201 - angle_012

 outAngle = abs(angle_201-60.0) +abs(angle_012-60.0) +abs(angle_123-60.0)
 GetElemDistortionSODA_Tria = True
End Function


'--------------------------------------------------------------------------


Sub StatusBar( Last As Long, optional femapMod as femap.model )
'
' - Description
'     This methode permit to show, hide or update progress bar of Femap
'
' - Input:
'    femapMod -> link to Femap application
'    last     -> >0 Show status bar and set max value
'                =0 Increment status bar
'                <0 Hide status bar
' - output:
'
' - Remarks/Usage:
'    Updates of status bar are automatically adjust for best results
'
  Static SCount As Long, UpdateStep As Long, inc as long, femapApp as femap.model

  If Last=0 Then    ' Incremente
     inc = inc +1
     If inc = UpdateStep Then
        inc=0 : SCount = SCount+ UpdateStep
        femapApp.feAppStatusUpdate( SCount  )
     End If
  ElseIf  Last> 0 Then ' ini
     set femapApp=femapMod : SCount = 0
     Call femapApp.feAppStatusShow( True, Last )
     UpdateStep= CLng(Last/50 + 0.49999999)
  Else   ' Hide
     Call femapApp.feAppStatusShow(False, 0)
     Set femapApp = nothing
  End If

End Sub


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

