Dim speech
Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")

msg = inputbox("i am gpt speaker","bot")
objXMLHTTP.Open "GET", "https://nahanabzar.ir/ai?text="+msg, False
objXMLHTTP.Send
stro = objXMLHTTP.ResponseText

set speech=CreateObject("sapi.spvoice")
speech.speak stro
