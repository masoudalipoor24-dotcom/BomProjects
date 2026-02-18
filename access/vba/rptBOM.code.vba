Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    Dim indent As Long
    indent = Nz(Me!LevelNo, 1) * 300

    Me!txtComponentCode.Left = 500 + indent
    Me!txtComponentDescription.Left = 1500 + indent
End Sub
