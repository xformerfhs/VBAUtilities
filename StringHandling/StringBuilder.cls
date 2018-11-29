VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Private constants for error messages
'
Private Const MODULNAME As String = "StringBuilder"

Private Const N_START_ERROR_MESSAGE As Long = vbObjectError + 27220

Private Const ERR_BLOCK_SIZE_TOO_SMALL As Long = N_START_ERROR_MESSAGE
Private Const STR_ERR_BLOCK_SIZE_TOO_SMALL As String = "Block size too small: "

Private Const ERR_BLOCK_SIZE_TOO_LARGE As Long = N_START_ERROR_MESSAGE + 1
Private Const STR_ERR_BLOCK_SIZE_TOO_LARGE As String = "Block size too large: "

Private Const ERR_TOO_LONG As Long = N_START_ERROR_MESSAGE + 2
Private Const STR_ERR_TOO_LONG As String = "Length exceeds maximum length for strings"

Private Const ERR_LENGTH_INCREASED As Long = N_START_ERROR_MESSAGE + 3
Private Const STR_ERR_LENGTH_INCREASED As String = "The length must not be increased without content"

Private Const ERR_LENGTH_LESS_THAN_ZERO As Long = N_START_ERROR_MESSAGE + 4
Private Const STR_ERR_LENGTH_LESS_THAN_ZERO As String = "Length must not be less than zero"

'
' Private constants
'
Private Const DEFAULT_BLOCK_SIZE As Long = 64& * 1024&
Private Const MINIMUM_BLOCK_SIZE As Long = 1024&
Private Const MAXIMUM_BLOCK_SIZE As Long = 64& * 1024& * 1024&

Private Const MAXIMUM_LENGTH As Long = &H7FFFFFFF

'
' Public constants
'
Public Enum TSBBlockSize
  sbbsMinimum  = MINIMUM_BLOCK_SIZE
  sbbsgMaximum = MAXIMUM_BLOCK_SIZE
End Enum

'
' Instance variables
'
Private m_Content As String
Private m_Length As Long
Private m_BlockSize As Long
Private m_Size As Long

'
' Public methods
'
Public Property Get blockSize() As Long
   blockSize = m_BlockSize
End Property

Public Property Let blockSize(ByVal newBlockSize As Long)
   If newBlockSize < MINIMUM_BLOCK_SIZE Then
      Err.Raise ERR_BLOCK_SIZE_TOO_SMALL, MODULNAME, STR_ERR_BLOCK_SIZE_TOO_SMALL & Format$(newBlockSize)
   Else
      If newBlockSize > MAXIMUM_BLOCK_SIZE Then
         Err.Raise ERR_BLOCK_SIZE_TOO_LARGE, MODULNAME, STR_ERR_BLOCK_SIZE_TOO_LARGE & Format$(newBlockSize)
      Else
         m_BlockSize = newBlockSize
      End If
   End If
End Property

Public Property Get content() As String
   content = Left$(m_Content, m_Length)
End Property

Public Property Let content(ByVal newContent As String)
   Me.reset
   Me.append newContent
End Property

Public Property Get length() As Long
   length = m_Length
End Property

Public Property Get isEmpty() As Boolean
   isEmpty = (m_Length = 0)
End Property

Public Property Let length(ByVal newLength As Long)
   If newLength > m_Length Then
      Err.Raise ERR_LENGTH_INCREASED, MODULNAME, STR_ERR_LENGTH_INCREASED
   Else
      If newLength < 0 Then
         Err.Raise ERR_LENGTH_LESS_THAN_ZERO, MODULNAME, STR_ERR_LENGTH_LESS_THAN_ZERO
      Else
         m_Length = newLength
      End If
   End If
End Property

Public Function setTo(ByVal text As String) As StringBuilder
   Me.reset
   
   Set setTo = Me.append(text)
End Function

Public Function append(ByVal text As String) As StringBuilder
   Dim additionalLength As Long
   
   additionalLength = Len(text)
   
   If additionalLength > 0 Then
      If (MAXIMUM_LENGTH - m_Length) < additionalLength Then
         Err.Raise ERR_TOO_LONG, MODULNAME, STR_ERR_TOO_LONG
      Else
         If (m_Length + additionalLength) > m_Size Then
            If (MAXIMUM_LENGTH - m_Length) < m_BlockSize Then
               m_BlockSize = MAXIMUM_LENGTH - m_Length
            End If
            
            m_Content = m_Content & Space$(m_BlockSize)
            m_Size = m_Size + m_BlockSize
         End If
      End If
   
      Mid$(m_Content, m_Length + 1, additionalLength) = text
   
      m_Length = m_Length + additionalLength
   End If

   Set append = Me
End Function

Public Function reset() As StringBuilder
   m_Length = 0
   
   Set reset = Me
End Function

'
' Class methods
'
Private Sub Class_Initialize()
   m_Length = 0
   m_BlockSize = DEFAULT_BLOCK_SIZE
   m_Size = 0
End Sub
