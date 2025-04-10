Function getSelectedText()
    Dim objMail As MailItem
    Dim objWordEditor As Word.Document
    Dim objSelection As Word.Selection
    
    ' Get the selected text from the active email
    Set objMail = Application.ActiveExplorer.Selection.Item(1)
    Set objWordEditor = objMail.GetInspector.WordEditor
    Set objSelection = objWordEditor.Application.Selection
    

    ' Check if there is any text selected
    If objSelection.Text = "" Then
        MsgBox "No text selected.", vbExclamation
        Exit Function

    End If
    
    ' Return the text within the range
    getSelectedText = RemoveNonPrintableCharsRegex(Replace(Left(objSelection.Text, Len(objSelection.Text) - 1), Chr(34), "`"))
    

End Function
Function RemoveNonPrintableCharsRegex(inputStr As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "[^\x20-\x7E]"  ' Match anything NOT in ASCII 32 (space) to 126 (~)
    regex.Global = True
    RemoveNonPrintableCharsRegex = regex.Replace(inputStr, "")
End Function
Sub get_email(string_var)
    ' Check if the OPENAI_API_KEY environment variable is set, and display an error message and exit the subroutine if it is not set
    If Environ("OPENAI_API_KEY") = "" Then
        MsgBox "Please set up your API key"
        Exit Sub
    End If
    new_var = get_res_from_OpenAIAPI(string_var)
    Call instert_msg_ToNextLine(new_var)

End Sub
Sub translate_My_email()
    
    ' Get the selected text from the active window, remove any line breaks, and store it in a variable named "test_var"
    test_var = Replace(Replace(getSelectedText, Chr(10), "\n"), Chr(13), "\n")
    ' If the selected text is not empty, prompt the user to review the text and write a professional email in English, then pass the text to the OpenAI API to generate a response
    If test_var <> "" Then
        propt = "review and translate my below email to Chinese, only reply the Chinese email you translated without any other input:"
        Call get_email(propt & test_var)
    End If

End Sub
Sub revise_My_email()

    ' Get the selected text from the active window, remove any line breaks, and store it in a variable named "test_var"
    test_var = Replace(Replace(getSelectedText, Chr(10), "\n"), Chr(13), "\n")
    ' If the selected text is not empty, prompt the user to review the text and write a professional email in English, then pass the text to the OpenAI API to generate a response
    If test_var <> "" Then
        propt = "review and revise my below email, only reply the email you revised without any other input:"
        Call get_email(propt & test_var)
    End If

End Sub
Sub Analyze_My_email()

    test_var = Replace(Replace(getSelectedText, Chr(10), "\n"), Chr(13), "\n")
    If test_var <> "" Then
        propt = "review and analyze below email, give me a easy-to-understand result, only reply the info in Markdown format. reply in Chinese but keep the necessary wordings in english."
        Call get_email(propt & test_var)
    End If

End Sub
Sub write_an_email()
        
    test_var = Replace(Replace(getSelectedText, Chr(10), "\n"), Chr(13), "\n")
    If test_var <> "" Then
        propt = "Review below info and write a professional email in English. Only reply to the email you write without any other input:"
        Call get_email(propt & test_var)
    End If
    
End Sub

Sub instert_msg_ToNextLine(new_var)
    Dim objMail As MailItem
    Dim objWordEditor As Word.Document
    Dim objRange As Word.Range
    
    ' Get the selected text from the active email
    Set objMail = Application.ActiveExplorer.Selection.Item(1)
    Set objWordEditor = objMail.GetInspector.WordEditor
    Set objRange = objWordEditor.Application.Selection.Range
    
    ' Move the cursor to the next line
    objRange.End = objRange.End
    objRange.Collapse wdCollapseEnd
    objRange.Select

    ' Get the value of the string variable
    strVariable = vbNewLine + new_var
    
    ' Insert the value of the string variable at the current position
    objWordEditor.Application.Selection.InsertAfter strVariable
End Sub

Function get_res_from_OpenAIAPI(str_p)
    ' Create an XMLHTTP request object to send the API request
    Dim request As Object
    Set request = CreateObject("MSXML2.XMLHTTP")
    
    ' Set the OpenAI API key
    apiKey = Environ("OPENAI_API_KEY")
 
    ' Set the URL of the OpenAI API endpoint
    Dim url As String
    'url = "https://api.openai.com/v1/chat/completions"
    url = "https://api.deepseek.com/v1/chat/completions"
    ' Get the data to send with the API request using the get_data function
    Dim data As String
    data = get_data(Trim(str_p))
    
    ' Send the API request using the XMLHTTP request object
    request.Open "POST", url, False
    request.setRequestHeader "Content-Type", "application/json"
    request.setRequestHeader "Authorization", "Bearer " & apiKey
    request.Send (data)

    If InStr(request.ResponseText, "You exceeded your current quota") Then
        MsgBox "Your OPENAI key has been expired, please change your OPENAI API key"
    Exit Function
    End If
    
    ' Get the response from the OpenAI API and parse the JSON using the handle_json function
    get_res_from_OpenAIAPI = Trim(handle_json(request.ResponseText, "query.choices[0].message.content"))
End Function
Function handle_json(json_string, json_para)
    ' Create a new ScriptControl object to handle JavaScript code

    Dim ScriptControl As Object
    Set ScriptControl = CreateObjectx86("MSScriptControl.ScriptControl")

    ScriptControl.Language = "JavaScript"
    ' Add the json string as a variable in the JavaScript environment
    ScriptControl.AddCode ("var query = " & json_string)
    ' Extract the response message from the JSON data and return it
    handle_json = ScriptControl.Eval(json_para)
End Function

Function get_data(prompt)
    ' Create a new Scripting.Dictionary object to store the API parameters
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    ' Add the necessary parameters to the dictionary
    dic.Add "model", "deepseek-chat"
    ' Use the prompt parameter to create a message in the required format
    dic.Add "messages", "[{" & Chr(34) & "role" & Chr(34) & ": " & Chr(34) & "user" & Chr(34) & "," & Chr(34) & "content" & Chr(34) & ": " & Chr(34) & prompt & Chr(34) & "}]"

    'dic.Add "temperature", 1.3
    'dic.Add "max_tokens", 1000
    ' Convert the dictionary to a JSON string and return it
 
    get_data = dic2json(dic)
End Function

Function dic2json(dic As Object)
    ' Initialize an empty string to store the JSON data
    msg = ""
    ' Get the keys and values from the dictionary
    k = dic.keys
    v = dic.Items
    ' Loop through the dictionary and convert each key-value pair to a JSON object property
    For i = 0 To dic.Count - 1
        Key = k(i)
        Value = v(i)
        ' If the value is an array, integer or double, don't add quotes around it
        If InStr(Value, "]") > 0 Or TypeName(Value) = "Integer" Or TypeName(Value) = "Double" Then
            msg = msg & Chr(34) & Key & Chr(34) & ":" & Value & ","
        ' Otherwise, add quotes around the value
        Else
            msg = msg & Chr(34) & Key & Chr(34) & ":" & Chr(34) & Value & Chr(34) & ","
        End If
    Next
    ' Remove the trailing comma and wrap the JSON data in curly braces
    msg = Left(msg, Len(msg) - 1)
    dic2json = "{" + msg + "}"
End Function

Function CreateObjectx86(Optional sProgID, Optional bClose = False)
Static oWnd As Object
Dim bRunning As Boolean
#If Win64 Then
bRunning = InStr(TypeName(oWnd), "HTMLWindow") > 0
If bClose Then
If bRunning Then oWnd.Close
Exit Function
End If
If Not bRunning Then
Set oWnd = CreateWindow()
oWnd.execScript "Function CreateObjectx86(sProgID): Set CreateObjectx86 = CreateObject(sProgID): End Function", "VBScript"
End If
Set CreateObjectx86 = oWnd.CreateObjectx86(sProgID)
#Else
Set CreateObjectx86 = CreateObject("MSScriptControl.ScriptControl")
#End If
End Function
Function CreateWindow()
Dim sSignature, oShellWnd, oProc
On Error Resume Next
 
    sSignature = Left(CreateObject("Scriptlet.TypeLib").GUID, 38)
    CreateObject("WScript.Shell").Run "%systemroot%\syswow64\mshta.exe about:""about:<head><script>moveTo(-32000,-32000);document.title='x86Host'</script><hta:application showintaskbar=no /><object id='shell' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shell.putproperty('" & sSignature & "',document.parentWindow);</script></head>""", 0, False
Do
 
For Each oShellWnd In CreateObject("Shell.Application").Windows
    Set CreateWindow = oShellWnd.GetProperty(sSignature)
    If Err.Number = 0 Then Exit Function
        Err.Clear
    Next
Loop
End Function
