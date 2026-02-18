Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(Cancel As Integer)
    Me!CreatedOn = Now()
    Me!CreatedBy = Environ$("USERNAME")
    Me!IsActive = True
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Me!ModifiedOn = Now()
    Me!ModifiedBy = Environ$("USERNAME")
End Sub
