VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'
' Private Konstanten für Fehlermeldungen
'
Private Const MODULNAME As String = "Stack"

Private Const N_START_ERROR_MESSAGE As Long = vbObjectError + 2600

Private Const ERR_STACK_IS_EMPTY     As Long = N_START_ERROR_MESSAGE
Private Const STR_ERR_STACK_IS_EMPTY As String = "Stack is empty"

'
' Private Konstanten
'
Private Const DEFAULT_BLOCKSIZE As Long = 50
Private Const MIN_BLOCKSIZE     As Long = 10
Private Const MAX_BLOCKSIZE     As Long = 100000

'
' Instanzvariable
'
Private m_Items() As Variant
Private m_BlockSize As Long
Private m_ActSize As Long
Private m_ActIndex As Long

'
' Private Methoden
'

'
' Öffentliche Methoden
'
Public Property Get blockSize() As Long
   blockSize = m_BlockSize
End Property

Public Property Let blockSize(ByVal newBlockSize As Long)
   If newBlockSize < MIN_BLOCKSIZE Then _
      newBlockSize = MIN_BLOCKSIZE

   If newBlockSize > MAX_BLOCKSIZE Then _
      newBlockSize = MAX_BLOCKSIZE

   m_BlockSize = newBlockSize
End Property

Public Property Get isEmpty() As Boolean
   isEmpty = (m_ActIndex = 0)
End Property

Public Property Get depth() As Long
   depth = m_ActIndex
End Property

Public Property Get topElement() As Variant
   If Me.isEmpty Then
      Err.Raise ERR_STACK_IS_EMPTY, MODULNAME, STR_ERR_STACK_IS_EMPTY
   Else
      topElement = m_Items(m_ActIndex)
   End If
End Property

Public Sub reset()
   m_ActIndex = 0
End Sub

Public Sub push(ByVal var As Variant)
   Dim newIndex As Long

   newIndex = m_ActIndex + 1

   If newIndex > m_ActSize Then
      m_ActSize = m_ActSize + m_BlockSize
      ReDim Preserve m_Items(1 To m_ActSize)
   End If

   m_Items(newIndex) = var

   m_ActIndex = newIndex
End Sub

Public Function pop() As Variant
   If Me.isEmpty Then
      Err.Raise ERR_STACK_IS_EMPTY, MODULNAME, STR_ERR_STACK_IS_EMPTY
   Else
      pop = m_Items(m_ActIndex)
      m_ActIndex = m_ActIndex - 1
   End If
End Function

Public Function dump(Optional ByVal delim As String = ", ") As String
   Dim dumpText As String
   Dim i As Long

   If Me.isEmpty Then
      dumpText = "<Empty>"
   Else
      For i = 1 To m_ActIndex
         dumpText = dumpText & delim & Format$(m_Items(i))
      Next i
   End If

   dump = Right$(dumpText, Len(dumpText) - Len(delim))
End Function

'
' Klassenmethoden
'
Private Sub Class_Initialize()
   m_BlockSize = DEFAULT_BLOCKSIZE
   m_ActSize = 0
   m_ActIndex = 0
End Sub
