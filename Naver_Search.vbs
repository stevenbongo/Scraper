Sub Naver_search()
    Set oShell = CreateObject("WScript.shell")
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Navigate "https://www.naver.com"
    objIE.Visible = True 

    ''' Wait until the html page is fully loaded '''
    Do Until objIE.ReadyState = 4
        WScript.Sleep 100
    Loop 

    ''' input text to search '''
    With objIE.Document
        .getElementByid("query").Value = "코로나바이러스"
        .getElementByid("search_btn").click()
    End With 

    Set a = objIE.Document.getElementsByClassName("container")
    a.click()
  
End Sub 


Call Naver_search()