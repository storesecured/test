<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->


<%

server.scripttimeout=3000
sDate=DateAdd("d",-1,now())
strYYYY = CStr(DatePart("yyyy", sDate))

strMM = CStr(DatePart("m", sDate))
If Len(strMM) = 1 Then strMM = "0" & strMM
strDD = DatePart("d", now())

sDate=strMM&strYYYY

on error resume next
sql_select = "select distinct store_id,site_name,service_type,store_email,bandwidth_subscription from store_settings where store_id<>101"
	set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

        Set fso = CreateObject("Scripting.FileSystemObject")

	if noRecords = 0 then
        FOR rowcounter= 0 TO myfields("rowcount")
                sStore_id = mydata(myfields("store_id"),rowcounter)
                site_name = mydata(myfields("site_name"),rowcounter)
                service_type = mydata(myfields("service_type"),rowcounter)
                sEmail = mydata(myfields("store_email"),rowcounter)
                bandwidth_subscription = mydata(myfields("bandwidth_subscription"),rowcounter)

                folderspec = fn_get_sites_folder(sStore_id,"Log")

                Set fsub = fso.GetFolder(folderspec)
                Set fc = fsub.Files
                For Each f2 in fc
                  if instr(f2,"stats"&sDate&".")>0 and instr(f2,".txt")>0 then
                    set t=fso.OpenTextFile(f2,1,false)
                    strFileContents = t.ReadAll
                    sPosStart = InStr(1,strFileContents,"BEGIN_DOMAIN",1)
                    sPosEnd = InStr(1,strFileContents,"END_DOMAIN",1)-14
                    sPosEnd = sPosEnd-sPosStart
                    If sPosStart>0 then
                       sStore_Id = replace(sName,"stores_","")
                       sBandwidth = mid(strFileContents,sPosStart+14,sPosEnd)
                       sTotal=0
                       for each sLine in split(sBandwidth,vbcrlf)

                           sLine = trim(sLine)
            

                           if sLine <> "" then
                              sBandwidthArray = split(sLine," ")
                              sSize= sBandwidthArray(3)
                              if isNumeric(sSize) then
                                 sTotal = sTotal + (sBandwidthArray(3)/1024)
                              end if
                           end if
            
                       next

                       sTotal = sTotal / 1024 / 1024

                       if sTotal > 1 then

                          sBandwidth=0
                          if service_type=3 or service_type=1 then
                             sBandwidth=1
                          elseif service_type=5 then
                             sBandwidth=2
                          elseif service_type=7 then
                             sBandwidth=4
                          elseif service_type=9 or service_type=10 then
                             sBandwidth=6
                          elseif service_type=11 then
                             sBandwidth=8
                          elseif service_type=12 then
                             sBandwidth=24
                          end if
                          
                          sBandwidth=sBandwidth+(bandwidth_subscription*5)
                          sOverrage=formatnumber(sTotal-sBandwidth,2)
                          sOverrageCharge= formatnumber(sOverrage*5,2)

                          if sTotal > sBandwidth and sOverrageCharge>1 then

                             if strDD="1" then
                               sSubject="StoreSecured Bandwidth Overage Notice"
                                'this is the actual final email
                               sText=vbcrlf&"Dear Store Owner," &_
                                  vbcrlf&vbcrlf&"Your StoreSecured store, http://"&site_name&" (#"&my_store_id&"), has exceeded its allowable bandwidth of "&sBandwidth&" GB per month."&_
                                  vbcrlf&vbcrlf&"Bandwidth overages are charged at the rate of $5 per GB."&_
                                  vbcrlf&vbcrlf&"Your current usage was "&formatnumber(sTotal,2)&" GB which is  "&sOverrage&" GB over your allowed usage for a total charge of $"&sOverrageCharge&_
                                  vbcrlf&vbcrlf&"This amount will be billed to your credit card within the next 24 hours."&_
                                  vbcrlf&vbcrlf&"If you do not understand what bandwidth is or wish to learn more about how to reduce your usage please visit the following URL http://server.iad.liveperson.net/hc/s-7400929/cmd/kbresource/kb-63579114663830375/view_question!PAGETYPE?sq=bandwidth&sf=101113&sg=0&st=66374&documentid=203362&action=view"
                               if sOverrageCharge>15 then
                                  sText=sText&vbcrlf&vbcrlf&"If you would like to upgrade your existing bandwidth subscription, subscriptions can be purchased at a reduced rate of $15 per 5GB additional bandwidth."&_
                                          vbcrlf&vbcrlf&"Since your overrage has exceeded the subscription charge of $15 we would recommend that if you plan to continue this level of usage that you purchase a bandwidth subscription to reduce your costs for future months.  "&_
                                          vbcrlf&vbcrlf&"To activate a bandwidth subscription please submit a support request with the subject Bandwidth Subscription and indicate how much of a subscription you would like to purchase, ie how many total GB."
                               end if
                             else
                               sSubject="StoreSecured Bandwidth Overage Total"
                               sText=vbcrlf&"Dear Store Owner," &_
                                  vbcrlf&vbcrlf&"Your StoreSecured store, http://"&site_name&" (#"&my_store_id&"), has exceeded its allowable bandwidth of "&sBandwidth&" GB per month."&_
                                  vbcrlf&vbcrlf&"Bandwidth overages are charged at the rate of $5 per GB."&_
                                  vbcrlf&vbcrlf&"Your current usage is "&formatnumber(sTotal,2)&" GB which is  "&sOverrage&" GB over your allowed usage for a total excess charge so far of $"&sOverrageCharge&_
                                  vbcrlf&vbcrlf&"The total overage charge will be billed to your credit card on the 1st of the month based on your month end bandwidth usage."&_
                                  vbcrlf&vbcrlf&"If you do not understand what bandwidth is or wish to learn more about how to reduce your usage please visit the following URL http://server.iad.liveperson.net/hc/s-7400929/cmd/kbresource/kb-63579114663830375/view_question!PAGETYPE?sq=bandwidth&sf=101113&sg=0&st=66374&documentid=203362&action=view"&_
                                  vbcrlf&vbcrlf&"This is a courtesy notice of the overage so that you may plan accordingly or try to lower your usage going forward.  Your store will NOT be closed due to this overage.  Please note that bandwidth is calculated based on the calendar month."

                               if sOverrageCharge>15 then
                                  sText=sText&vbcrlf&vbcrlf&"If you would like to upgrade your existing bandwidth subscription, subscriptions can be purchased at a reduced rate of $15 per 5GB additional bandwidth."&_
                                          vbcrlf&vbcrlf&"Since your overrage has exceeded the subscription charge of $15 we would recommend that if you plan to continue this level of usage that you purchase a bandwidth subscription to reduce your costs for future months.  "&_
                                          vbcrlf&vbcrlf&"To activate a bandwidth subscription for next month please submit a support request with the subject Bandwidth Subscription and indicate how much of a subscription you would like to purchase, ie how many total GB."
                               end if

                             end if

                             sText=sText&vbcrlf&vbcrlf&"Sincerely,"&vbcrlf&vbcrlf&"StoreSecured Support"

                             Send_Mail sNoReply_email,sEmail,"StoreSecured Bandwidth Overage Notification",sText

                             sTextTotal=sTextTotal&vbcrlf&my_store_id&chr(9)&sOverrageCharge&chr(9)&sEmail

                          end if
                          if bandwidth_subscription>0 then
                                sTextTotal=sTextTotal&vbcrlf&my_store_id&chr(9)&bandwidth_subscription*15&chr(9)&sEmail&chr(9)&"sub"
                          end if
                       end if

                    end if
                    t.close
                  end if
              next

        Next
	End If
	
	set fso=nothing

Send_Mail sReport_email,sReport_email,"Bandwidth Charges",sTextTotal



%>
