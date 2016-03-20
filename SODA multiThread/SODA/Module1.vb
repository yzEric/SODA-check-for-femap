'______________________________________________________________
'    Name:       SODA - Sum Of Deviation Angles
'    Author:     E. LE GAL
'    Version:    1.0
'    Languages:  Visual Basic
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

Option Explicit On    '-> All variables must be declared




Module Module1

    Private PI As Double
    Private halfPI As Double
    Private radToDeg_coeff As Double


    Sub ini_const()
        PI = 4 * Math.Atan(1)
        halfPI = 2 * Math.Atan(1)
        radToDeg_coeff = 45 / Math.Atan(1)
    End Sub



    Sub Main()
        Call ini_const()

        Dim rc As femap.zReturnCode
        Dim femapMod As femap.model = GetObject(, "femap.model")


        '--- Aks User to select elements
        Dim elemSet As femap.Set = femapMod.feSet
        rc = elemSet.Select(femap.zDataType.FT_ELEM, True, "Select 2D elements")
        If AssertRC(femapMod, rc, "Unable to select elements") Then Exit Sub

        Dim UserSelectCount As Integer = elemSet.Count

        '--- Start crono
        Dim startTime As Double = Timer


        '--- Keep only 2D shape elements
        Dim shapeSet As femap.Set = femapMod.feSet
        Dim All2DElemsSet As femap.Set = femapMod.feSet

        rc = shapeSet.AddArray(4, {femap.zTopologyType.FTO_TRIA3, femap.zTopologyType.FTO_TRIA6, femap.zTopologyType.FTO_QUAD4, femap.zTopologyType.FTO_QUAD8})
        If AssertRC(femapMod, rc, "Unable to add data to shapeSet") Then Exit Sub

        rc = All2DElemsSet.AddSetRule(shapeSet.ID, femap.zGroupDefinitionType.FGD_ELEM_BYSHAPE)
        If AssertRC(femapMod, rc, "Unable to create a set of element by using shapes") Then Exit Sub

        rc = elemSet.RemoveNotCommon(All2DElemsSet.ID)
        If AssertRC(femapMod, rc, "Unable to exclude non 2D elements") Then Exit Sub

        If UserSelectCount <> elemSet.Count Then femapMod.feAppMessage(femap.zMessageColor.FCM_HIGHLIGHT, CStr(elemSet.Count) + " elements removed from selection")

        shapeSet = Nothing
        All2DElemsSet = Nothing

        If elemSet.Count = 0 Then
            femapMod.feAppMessage(femap.zMessageColor.FCM_HIGHLIGHT, "No selected element can be used")
            Exit Sub
        End If


        '--- Get Coordinates of nodes used by elements
        Call femapMod.feAppMessage(femap.zMessageColor.FCM_NORMAL, "1/2 : Get coordinate of nodes")

        Dim nodeSet As femap.Set = femapMod.feSet
        Dim coordTable(,) As Double = {}

        rc = nodeSet.AddSetRule(elemSet.ID, femap.zGroupDefinitionType.FGD_NODE_ONELEM)
        If AssertRC(femapMod, rc, "Uanble to get list of nodes used by elements") Then Exit Sub

        If Not GetNodesCoord_bigTable(femapMod, nodeSet, coordTable) Then Exit Sub

        nodeSet = Nothing



        '--- Calculation of distortions
        Call femapMod.feAppMessage(femap.zMessageColor.FCM_NORMAL, "2/2 : Calculation of criteria")

        Dim SODA() As Double = {}
        Dim elemIDs() As Integer = {}
        If Not getDistortionSODA(femapMod, elemSet, coordTable, True, SODA, elemIDs) Then Exit Sub

        elemSet = Nothing


        '--- Create outputSet
        Const outputSet_Title = "Distortion SODA"
        Dim OutputSetID As Integer = femapMod.Info_NextID(femap.zDataType.FT_OUT_CASE)

        If Not createOutputSet(femapMod, OutputSetID, outputSet_Title, "Sum of deviation angles", 0) Then Exit Sub


        '--- Write results ---
        Const vectorID = 400000 ' 9000000
        Const vectorTitle = "Sum of deviation angle"

        Dim outVector As femap.Output = femapMod.feOutput
        Dim count As Integer = UBound(elemIDs) + 1


        rc = outVector.InitScalarAtElem(OutputSetID, vectorID, vectorTitle, 10, False)
        If AssertRC(femapMod, rc, "Unable to initialize the output vector") Then Exit Sub

        Dim v1 As Object = elemIDs, v2 As Object = SODA

        rc = outVector.PutScalarAtElem(count, elemIDs, SODA)
        If AssertRC(femapMod, rc, "Unable to put results in the output vector") Then Exit Sub

        rc = outVector.Put(0)
        If AssertRC(femapMod, rc, "Unable to put the output vector in the model") Then Exit Sub

        Call femapMod.feAppMessage(femap.zMessageColor.FCM_NORMAL, "Output Set " + CStr(OutputSetID) + ": " + CStr(vectorID) + ".." + vectorTitle)


        '--- This is The End
        femapMod.feAppMessage(femap.zMessageColor.FCM_NORMAL, "--- Done in " + Format((Timer - startTime), "0.0s - End of task"))
        femapMod = Nothing
    End Sub




    Function GetNodesCoord_bigTable(ByRef femapMod As femap.model, ByRef nodeSet As femap.Set, ByRef outTable(,) As Double) As Boolean
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

        If nodeSet.Count = 0 Then
            GetNodesCoord_bigTable = True : Exit Function
        End If

        Dim nodeCount As Integer, xyz As Object, IDs As Object, rc As femap.zReturnCode
        Dim aNode As femap.Node = femapMod.feNode

        rc = aNode.GetCoordArray(nodeSet.ID, nodeCount, IDs, xyz)
        If AssertRC(femapMod, rc, "Unable to obtain coordinates of nodes") Then Return False

        Dim maxID As Integer = nodeSet.Last
        ReDim outTable(2, maxID)

        Dim nodeID As Integer, elemIndex As Integer
        Dim nodes_Last As Integer = nodeCount - 1

        For i As Integer = 0 To nodes_Last
            nodeID = IDs(i)
            outTable(0, nodeID) = xyz(elemIndex)
            outTable(1, nodeID) = xyz(elemIndex + 1)
            outTable(2, nodeID) = xyz(elemIndex + 2)
            elemIndex = elemIndex + 3
        Next i

        Return True
    End Function



    Function GetElemDistortionSODA_Quad( _
     ByVal x0 As Double, ByVal y0 As Double, ByVal z0 As Double, _
     ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, _
     ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, _
     ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double) As Double
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
        Dim angle_301 As Double = AngleOf2Vectors_deg(x3 - x0, y3 - y0, z3 - z0, x1 - x0, y1 - y0, z1 - z0)
        Dim angle_012 As Double = AngleOf2Vectors_deg(x0 - x1, y0 - y1, z0 - z1, x2 - x1, y2 - y1, z2 - z1)
        Dim angle_123 As Double = AngleOf2Vectors_deg(x1 - x2, y1 - y2, z1 - z2, x3 - x2, y3 - y2, z3 - z2)
        Dim angle_230 As Double = 360.0 - angle_301 - angle_012 - angle_123

        Return Math.Abs(angle_301 - 90.0) + Math.Abs(angle_012 - 90.0) + Math.Abs(angle_123 - 90.0) + Math.Abs(angle_230 - 90.0)
    End Function



    Function GetElemDistortionSODA_Tria( _
     ByVal x0 As Double, ByVal y0 As Double, ByVal z0 As Double, _
     ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, _
     ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double) As Double
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
        Dim angle_201 As Double = AngleOf2Vectors_deg(x2 - x0, y2 - y0, z2 - z0, x1 - x0, y1 - y0, z1 - z0)
        Dim angle_012 As Double = AngleOf2Vectors_deg(x0 - x1, y0 - y1, z0 - z1, x2 - x1, y2 - y1, z2 - z1)
        Dim angle_123 As Double = 180.0 - angle_201 - angle_012

        Return Math.Abs(angle_201 - 60.0) + Math.Abs(angle_012 - 60.0) + Math.Abs(angle_123 - 60.0)
    End Function



    Function getDistortionSODA(ByRef femapMod As femap.model, ByRef elemSet As femap.Set, ByRef coordTable As Object, ByVal showStatusBar As Boolean, ByRef outSODA() As Double, ByRef outIDs() As Integer) As Boolean
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

        If elemSet.Count = 0 Then
            Return True
        End If


        '--- Get data of elements : nodes
        Dim rc As femap.zReturnCode
        Dim anElem As femap.Elem = femapMod.feElem

        Dim numElem As Integer, entID As Object, propID As Object, elemTYPE As Object, topology As Object, layerID As Object
        Dim color As Object, formulation As Object, orient As Object, offset As Object, release As Object
        Dim orientSET As Object, orientID As Object, ElemNodes As Object, connectTYPE As Object, connectSEG As Object

        rc = anElem.GetAllArray(elemSet.ID, numElem, entID, propID, elemTYPE, topology, layerID, color, formulation, orient, offset, release, orientSET, orientID, ElemNodes, connectTYPE, connectSEG)
        If AssertRC(femapMod, rc, "Unable to obtain data of elements") Then Return False

        Erase propID, elemTYPE, layerID, color, formulation, orient, offset, release, orientSET, orientID, connectTYPE, connectSEG
        'keep: entID, topology, ElemNodes 



        '--- Compute SODA

        Dim lastElem As Integer = numElem - 1
        ReDim outSODA(lastElem), outIDs(lastElem)

        For i_Elem As Integer = 0 To lastElem
            outIDs(i_Elem) = entID(i_Elem)
        Next

        Dim SODAdata As New Thread_SODA
        SODAdata.coordTable = coordTable
        SODAdata.ElemNodes = ElemNodes
        SODAdata.topology = topology
        SODAdata.outSODA = outSODA


        Dim thCount As Integer = System.Environment.ProcessorCount
        femapMod.feAppMessage(femap.zMessageColor.FCM_HIGHLIGHT, "Nb CPU: " + CStr(thCount))

        Dim th() As System.Threading.Thread, Params As Object, aThread As System.Threading.Thread
        ReDim th(thCount - 1)
        ReDim Params(thCount - 1)

        Dim startPos As Integer, endPos As Integer
        Dim size As Integer = Math.Ceiling((lastElem + 1) / thCount)
        endPos = -1

        For iTh As Integer = 0 To thCount - 1
            startPos = endPos + 1
            endPos = Math.Min(startPos + size, lastElem)

            Params = {startPos, endPos}

            aThread = New System.Threading.Thread(AddressOf SODAdata.SODA_multiThread)
            th(iTh) = aThread

            aThread.Start(Params)
        Next


        For iTh As Integer = 0 To thCount - 1
            th(iTh).Join()
        Next



        Return True
    End Function









    Function AngleOf2Vectors_deg( _
              ByVal Xa As Double, ByVal Ya As Double, ByVal Za As Double, _
              ByVal Xb As Double, ByVal Yb As Double, ByVal Zb As Double) As Double
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
        '   cos angle = (Xa.Xb+Ya.Yb+Za.Zb) / sqrt((Xa^2+Ya^2+Za^2)(Xb^2+Yb^2+Zb^2 ))

        On Error GoTo errorInOp ' It is faster to catch error divBy0 Than to prevent it 

        Return radToDeg_coeff * Math.Acos((Xa * Xb + Ya * Yb + Za * Zb) / Math.Sqrt((Xa ^ 2 + Ya ^ 2 + Za ^ 2) * (Xb ^ 2 + Yb ^ 2 + Zb ^ 2)))


errorInOp:
        If (Err.Number = 10061) Then
            Err.Clear()
            AngleOf2Vectors_deg = 0
        Else
            Err.raise(Err.Number)
        End If

    End Function





    Function createOutputSet(ByVal femapMod As femap.model, ByVal ID As Integer, ByVal Title As String, ByVal notes As String, ByVal value As Double) As Boolean
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
        Dim rc As femap.zReturnCode
        Dim anOutputSet As femap.OutputSet = femapMod.feOutputSet

        anOutputSet.title = Title
        anOutputSet.notes = notes
        anOutputSet.Value = value
        anOutputSet.program = femap.zAnalysisProgram.FAP_FEMAP_GEN
        anOutputSet.analysis = femap.zAnalysisType.FAT_UNKNOWN

        rc = anOutputSet.Put(ID)
        If AssertRC(femapMod, rc, "Unable to create the OutputSet") Then Return False

        createOutputSet = True
    End Function



    Function AssertRC(ByVal femapMod As femap.model, ByVal rc As femap.zReturnCode, ByVal msg As String) As Boolean
        If rc = femap.zReturnCode.FE_OK Then Return False

        Dim info As String
        Select Case rc
            Case femap.zReturnCode.FE_TOO_SMALL
                info = "Too small"
            Case femap.zReturnCode.FE_FAIL
                info = "Fail"
            Case femap.zReturnCode.FE_BAD_TYPE
                info = "Bad type"
            Case femap.zReturnCode.FE_CANCEL
                info = "Cancel"
            Case femap.zReturnCode.FE_BAD_DATA
                info = "Bad data"
            Case femap.zReturnCode.FE_INVALID
                info = "Invalid"
            Case femap.zReturnCode.FE_NO_MEMORY
                info = "No memory"
            Case femap.zReturnCode.FE_NOT_EXIST
                info = "Not Exist"
            Case femap.zReturnCode.FE_NEGATIVE_MASS_VOLUME
                info = "Negative mass volume"
            Case femap.zReturnCode.FE_SECURITY
                info = "Security"
            Case femap.zReturnCode.FE_NO_FILENAME
                info = "No file name"
            Case femap.zReturnCode.FE_NOT_AVAILABLE
                info = "Not available"
            Case Else
                info = "Unkonw return code: " + CStr(rc)
        End Select

        info = "Return code is '" + info + "'"

        On Error Resume Next

        femapMod.feAppMessage(femap.zMessageColor.FCM_ERROR, msg)
        femapMod.feAppMessage(femap.zMessageColor.FCM_ERROR, info)

        MsgBox(msg + vbCrLf + info)
        Return True
    End Function



End Module
