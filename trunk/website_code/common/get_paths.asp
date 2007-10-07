<%

sIP = request.servervariables("Local_Addr")

select case sIP
       case "10.235.158.132","192.168.1.20","192.168.1.161","192.168.1.160"
            ServerName="S85789"
            sCol="content_20_done"
        case "10.235.158.133","192.168.1.27"
            ServerName="96036-app4"
            sCol="content_27_done"
        case "10.235.158.134","192.168.1.28"
            ServerName="96039-app5"
            sCol="content_28_done"
       case else
            'most likely a test server here so use user configured values
            ServerName=sComputerName
            sCol="content_28_done"
end select

strServerPort = Request.ServerVariables("SERVER_PORT")

freeserver=10
paidserver=9

%>
