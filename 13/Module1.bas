Attribute VB_Name = "modEtiket"
Public Sub EtiketEkle(Etiket$, Dur As Boolean)
Dim S As Long
    S = frmAna.rchHtml.SelStart
    If Len(frmAna.rchHtml.SelText) > 0 Then frmAna.rchHtml.SelText = ""
    frmAna.rchHtml.SelText = Etiket$
End Sub

