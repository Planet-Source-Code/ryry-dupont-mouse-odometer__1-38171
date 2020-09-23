Attribute VB_Name = "basRounder"
'######################################################
'##                                                  ##
'##            This code is From                     ##
'##                                                  ##
'##                                                  ##
'##            www.AbrstactVB.com                    ##
'##                                                  ##
'##                                                  ##
'##                                                  ##
'##                                                  ##
'##                                                  ##
'######################################################

Declare Function CreateRoundRectRgn Lib "gdi32" _
        (ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long) As Long
        
Declare Function SetWindowRgn Lib "user32" _
        (ByVal hwnd As Long, ByVal hRgn As Long, _
        ByVal bRedraw As Boolean) As Long


Public Sub MakeRound(pForm As Form, lValue As Long)
    Dim lRet As Long
    Dim l As Long
    Dim llWidth As Long
    Dim llHeight As Long
            
    'Get Form size in pixels
    llWidth = pForm.Width / Screen.TwipsPerPixelX
    llHeight = pForm.Height / Screen.TwipsPerPixelY
    
    'Create Form with Rounded Corners
    lRet = CreateRoundRectRgn(0, 0, llWidth, llHeight, _
                              lValue, lValue)
                              
    l = SetWindowRgn(pForm.hwnd, lRet, True)
End Sub


