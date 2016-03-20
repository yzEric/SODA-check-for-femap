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




Public Class Thread_SODA

    Public coordTable As Object
    Public ElemNodes As Object
    Public topology As Object
    Public outSODA() As Double



    Sub SODA_multiThread(ByVal StartEndParam As Object)

        Dim n0 As Integer, n1 As Integer, n2 As Integer, n3 As Integer
        Dim startIndex As Integer = StartEndParam(0)
        Dim endIndex As Integer = StartEndParam(1)
        Dim elemIndex As Integer = startIndex * 20

        For i_elem As Integer = startIndex To endIndex

            Select Case topology(i_elem)
                Case femap.zTopologyType.FTO_QUAD4, femap.zTopologyType.FTO_QUAD8
                    n0 = ElemNodes(elemIndex)
                    n1 = ElemNodes(elemIndex + 1)
                    n2 = ElemNodes(elemIndex + 2)
                    n3 = ElemNodes(elemIndex + 3)
                    outSODA(i_elem) = GetElemDistortionSODA_Quad( _
                        coordTable(0, n0), coordTable(1, n0), coordTable(2, n0), _
                        coordTable(0, n1), coordTable(1, n1), coordTable(2, n1), _
                        coordTable(0, n2), coordTable(1, n2), coordTable(2, n2), _
                        coordTable(0, n3), coordTable(1, n3), coordTable(2, n3))

                Case femap.zTopologyType.FTO_TRIA3, femap.zTopologyType.FTO_TRIA6
                    n0 = ElemNodes(elemIndex)
                    n1 = ElemNodes(elemIndex + 1)
                    n2 = ElemNodes(elemIndex + 2)
                    outSODA(i_elem) = GetElemDistortionSODA_Tria( _
                        coordTable(0, n0), coordTable(1, n0), coordTable(2, n0), _
                        coordTable(0, n1), coordTable(1, n1), coordTable(2, n1), _
                        coordTable(0, n2), coordTable(1, n2), coordTable(2, n2))

                Case Else
                    'outSODA(i_Elem)=0

            End Select

            elemIndex = elemIndex + 20
        Next i_elem

    End Sub





End Class
