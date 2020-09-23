VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public FileName             As String
Public Width                As Long
Public Height               As Long
Public Depth                As Long
Public TypeOfImage          As IMGType
Private Const BufferSize    As Long = 65535
Public Enum IMGType
    UNKNOWN = 0
    GIF = 1
    JPEG = 2
    PNG = 3
    BMP = 4
End Enum
#If False Then
Private UNKNOWN, GIF, JPEG, PNG, BMP
#End If

Private Const MAX_PATH      As Integer = 260

Private Sub Class_Initialize()

    Height = 0
    Width = 0
    Depth = 0
    TypeOfImage = UNKNOWN

End Sub

Public Sub GetImageFileInfo()


Dim arrTemp(BufferSize) As Byte
Dim FileNumber          As Integer
Dim lngStep             As Long

    On Error GoTo ErrorSnap
    FileNumber = FreeFile()
    Open FileName For Binary As FileNumber
    Get #FileNumber, 1, arrTemp()
    Close FileNumber
    If arrTemp(0) = 137 Then
        If arrTemp(1) = 80 Then
            If arrTemp(2) = 78 Then
              
                TypeOfImage = PNG
                Select Case arrTemp(25)
                Case 0 ' Greyscale
                    Depth = arrTemp(24)
                Case 2 ' RGB Encoded
                    Depth = arrTemp(24) * 3
                Case 3 ' 8 bpp
                    Depth = 8
                Case 4 ' Greyscale with alpha
                    Depth = arrTemp(24) * 2
                Case 6 ' RGB encoded with alpha
                    Depth = arrTemp(24) * 4
                Case Else
                    TypeOfImage = UNKNOWN
                End Select
                If TypeOfImage Then
                    ' width of image
                    Width = arrTemp(19) + arrTemp(18) * 256
                    ' heght of image
                    Height = arrTemp(23) + arrTemp(22) * 256
                End If
            End If
        End If
        
    End If
   
    If arrTemp(0) = 71 Then
        If arrTemp(1) = 73 Then
            If arrTemp(2) = 70 Then
              
                TypeOfImage = GIF
                Width = arrTemp(6) + arrTemp(7) * 256
                Height = arrTemp(8) + arrTemp(9) * 256
                Depth = (arrTemp(10) And 7) + 1
            End If
        End If
        
    End If
    
    If arrTemp(0) = 66 Then
        If arrTemp(1) = 77 Then
            
            TypeOfImage = BMP
            Width = arrTemp(18) + arrTemp(19) * 256
            Height = arrTemp(22) + arrTemp(23) * 256
            Depth = arrTemp(28)
        End If
    End If
   
    If TypeOfImage = UNKNOWN Then
        Do
            If arrTemp(lngStep = &HFF And arrTemp(lngStep + 1) = &HD8 And arrTemp(lngStep + 2) = &HFF) Or (lngStep >= BufferSize - 10) Then
               Exit Do
            End If
            lngStep = lngStep + 1
        Loop
        lngStep = lngStep + 2
        If lngStep >= BufferSize - 10 Then
            Exit Sub
        End If
        Do
            Do
                If arrTemp(lngStep) = &HFF Then
                    If arrTemp(lngStep + 1) <> &HFF Then
                       Exit Do
                    End If
                End If
               
                lngStep = lngStep + 1
                If lngStep >= BufferSize - 10 Then
                    Exit Sub
                End If
            Loop
            lngStep = lngStep + 1
            Select Case arrTemp(lngStep)
            Case &HC0 To &HC3, &HC5 To &HC7, &HC9 To &HCB, &HCD To &HCF
                Exit Do
            End Select
            lngStep = lngStep + (arrTemp(lngStep + 2) + arrTemp(lngStep + 1) * 256)
            If lngStep >= BufferSize - 10 Then
                Exit Sub
            End If
        Loop
        TypeOfImage = JPEG
        Height = arrTemp(lngStep + 5) + arrTemp(lngStep + 4) * 256
        Width = arrTemp(lngStep + 7) + arrTemp(lngStep + 6) * 256
        Depth = arrTemp(lngStep + 8) * 8
    End If

Exit Sub

ErrorSnap:
    MsgBox " Error occurs : " & Err.Description, vbCritical

End Sub


