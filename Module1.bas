Attribute VB_Name = "Module1"
Declare Function AlphaBlending Lib "Alphablending.dll" _
            (ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest As Long, _
            ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, _
            ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal AlphaSource As Long) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Sub Blend(Destination As Object, Source As Object, Amount As Integer, X, Y, X2, Y2)
AlphaBlending Destination.hdc, X, Y, X2, Y2, Source.hdc, X, Y, X2, Y2, Amount
End Sub

