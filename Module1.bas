Attribute VB_Name = "Module1"
Option Explicit

Private Sub Submit_Click()
    Dim userInput As String
    Dim chatbotResponse As String
    
    ' Get user input from cell A2
    userInput = Range("A2").Value
    
    ' Call OpenAI API to get chatbot response
    chatbotResponse = GetOpenAIResponse(userInput)
    
    ' Display chatbot response in cell B2
    Range("B2").Value = chatbotResponse
End Sub

Function GetOpenAIResponse(userInput As String) As String
    Dim objHTTP As Object
    Dim apiKey As String
    Dim apiUrl As String
    Dim jsonBody As String
    Dim jsonResponse As String
    Dim json As Object
    
    ' Set your OpenAI API key here
    apiKey = "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"
    
    ' Set OpenAI API endpoint URL
    apiUrl = "https://api.openai.com/v1/chat/completions"
    
    ' Create JSON request body
    jsonBody = "{""messages"": [{""role"": ""system"", ""content"": ""You are a helpful assistant.""}, {""role"": ""user"", ""content"": """ & userInput & """}], ""max_tokens"": 50, ""n"": 1, ""stop"": ""\n"", ""temperature"": 0.7}"
    
    ' Create an HTTP object
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Send the request to OpenAI API
    With objHTTP
        .Open "POST", apiUrl, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .send jsonBody
        jsonResponse = .responseText
    End With
    
    ' Create a JSON parser object
    Set json = CreateObject("MSXML2.DOMDocument.6.0")
    json.LoadXML jsonResponse
    
    ' Extract chatbot response from JSON and return it
    GetOpenAIResponse = json.SelectSingleNode("//content").Text
End Function
