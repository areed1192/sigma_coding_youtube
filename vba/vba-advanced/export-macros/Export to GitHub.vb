Sub SaveVbaScriptsToGitHub()

'Declare variables related to the URL.
Dim username As String
Dim repo_name As String
Dim file_name As String
Dim access_token As String
Dim payload As String

'Declare variables related to the HTTP Request.
Dim xml_obj As MSXML2.XMLHTTP60

'Declare variables related to the Visual Basic Editor.
Dim VBAEditor As VBIDE.VBE
Dim VBProj As VBIDE.VBProject
Dim VBCodeMod As VBIDE.CodeModule
Dim VBRawCode As String

'Create a reference to the editor itself, MAKE SURE YOU HAVE MACRO SECURITY TURNED OFF.
Set VBAEditor = Application.VBE

'Create a reference to the projects in the Visual Basic Editor. This will return my Personal Macro Workbook.
Set VBProj = VBAEditor.VBProjects(1)

'Reference a single component and grab the Code Module.
Set VBCodeMod = VBProj.VBComponents.Item("General_Format").CodeModule

'Grab the Raw Code.
VBRawCode = VBCodeMod.Lines(StartLine:=1, Count:=VBCodeMod.CountOfLines)

'Base64 Encode the String.
RawCodeEncoded = EncodeBase64(text:=VBRawCode)

'Create a reference to the Microsoft XML library
Set xml_obj = New MSXML2.XMLHTTP60

    'Define URL Components
    base_url = "https://api.github.com/repos/"
    repo_name = "sigma_coding_youtube/"
    username = "areed1192/"
    file_name = "test.vb"
    access_token = "TOKEN"
    
    'Build the full URL.
    full_url = base_url + username + repo_name + "contents/" + file_name + "?ref=master"
    
    'HERE IS AN EXAMPLE OF A URL THAT I'M TRYING TO BUILD.
    '"https://api.github.com/repos/areed1192/sigma_coding_youtube/contents/test.txt?ref=master"
    
    'Open a new request, specify the method and the URL and make sure it's async.
    xml_obj.Open bstrMethod:="PUT", bstrURL:=full_url, varAsync:=True
    
    'Set the headers.
    xml_obj.setRequestHeader bstrHeader:="Accept", bstrValue:="application/vnd.github.v3+json"
    xml_obj.setRequestHeader bstrHeader:="Authorization", bstrValue:="token " + access_token

    'Define the Payload.
    payload = "{""message"":""This is my message"",""content"":"""
    payload = payload + Application.Clean(RawCodeEncoded)
    payload = payload + """}"
    
    'IDEALLY IT WILL LOOK LIKE THIS, JUST LONGER.
    'payload = "{""message"":""This is my message"",""content"":""IlN1YiBGb3JtYXRJb==""}"
        
    'Send the request.
    xml_obj.send varBody:=payload
    
    'Wait till the request is made.
    While xml_obj.readyState <> 4
        DoEvents
    Wend
    
    'Print out some info.
    Debug.Print "FULL URL: " + full_url
    Debug.Print "STATUS TEXT :" + xml_obj.statusText
    Debug.Print "PAYLOAD: " + payload
    
End Sub

Sub GrabVBAModule()

Dim VBAEditor As VBIDE.VBE
Dim VBProj As VBIDE.VBProject
Dim VBCodeMod As VBIDE.CodeModule
Dim VBRawCode As String

'Create a reference to the editor itself
Set VBAEditor = Application.VBE

'Create a reference to the projects in the Visual Basic Editor
Set VBProj = VBAEditor.VBProjects(1)

'Reference a single project and print some information about that module.
Set VBCodeMod = VBProj.VBComponents.Item("General_Format").CodeModule

'Grab the Raw Code.
VBRawCode = VBCodeMod.Lines(StartLine:=1, Count:=VBCodeMod.CountOfLines)

'Base64 Encode the String.
RawCodeEncoded = EncodeBase64(text:=VBRawCode)

'Print it out.
Debug.Print "Here is the Encoded Content: " + RawCodeEncoded

End Sub


Function EncodeBase64(text As String) As String
    
    'Define the variables.
    Dim arrData() As Byte
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
    
    'Convert the String to a Unicode String.
    arrData = StrConv(text, vbFromUnicode)
    
    'Define the DOM Document and b64 Node.
    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")
    
    'Define the DataType.
    objNode.DataType = "bin.base64"

    'Assign the Node Vaule.
    objNode.nodeTypedValue = arrData
    
    'Return the Encoded Text.
    EncodeBase64 = Replace(objNode.text, vbLf, "")
    
    Set objNode = Nothing
    Set objXML = Nothing
    
End Function

