<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../common/conn-utf.asp"-->
<!--#include file="../common/Function-utf.asp"-->
<!--#include file="../common/safe.asp"-->
<%Response.ContentType = "text/json"%>
<%
'插入会员收藏书本，传入书本ID，读取session 中的userID
	dim lngUserId, strArrFavId, i, lngInfoId, strArrInfoid, isShop
	dim rsFav, strsql, lngstate, strMsg, strTempSql, strUserTempCounts, strmaxuseCountsTemp, lngtempCounts 
	
	lngUserId = ConvertLong(Session("UserId") & "")
	strArrFavId = trim(request("FavId") & "")
	if strArrFavId = "" then strArrFavId = "0"
	if strArrInfoid = "" then strArrInfoid = "0"
	
	Set rsFav = Server.CreateObject("ADODB.RecordSet")
	strsql = "select * from Fav_t where UserId = " & lngUserId & " and FavId in (" & strArrFavId & ")"
	'response.write strsql
	if rsFav.state = 1 then rs.close
	rsFav.open strsql,conn,1,3
	if not(rsFav.bof or rsFav.eof) then
		i=0
		do while not(rsFav.bof or rsFav.eof)
			lngInfoId = rsFav("InfoId")
			strArrInfoid = strArrInfoid & "," & lngInfoId
			i=i+1
			rsFav.movenext
		loop
		'会员单次最大借书量
		strUserTempCounts = getUserInfo("useCounts")
		'会员己经借书数量
		strmaxuseCountsTemp = getUserShopCount()
		'判断书本当前库存是否有数量
		isShop = getusershopisCount(strArrInfoid,i)
		'获取会员总借书次数
		isshop2 = getUserInfo("maxuseCountsTemp")
		isshop2 = ConvertLong(isshop2 & "")
		
		if isShop = 0 then
			lngstate = 1
			strMsg = "您选中的书本中有库存量为0的书本，请返回刷新后再试。"
		else
			'i 是本次要提交的借书数量
			thisUserCount = strUserTempCounts-strmaxuseCountsTemp
			if (i>thisUserCount) then
				lngstate = 1
				strMsg = "超出单次最大借书量，您本次还可以再借【 " & thisUserCount & " 】本，请返回刷新后再试。"
			else
				if isshop2 > 0 then
					'扣总最借书次数
					strTempSql = "update user_t set maxuseCountsTemp = maxuseCountsTemp-1 where userid = " & lngUserId
					conn.Execute strTempSql
					
					dim TempArr, TempArr_len, thisInfoId
					TempArr = Split(strArrInfoid, ",")
					'response.write strArrInfoid
					For TempArr_len = 0 To UBound(TempArr)
						thisInfoId = ConvertLong(TempArr(TempArr_len) & "")
						
						if thisInfoId <> 0 then
							'扣库存
							strTempSql = "Update info_t Set iCount = iCount -1 Where infoId = " & thisInfoId
							'response.write strtempsql
							conn.Execute strTempSql
							'创建借书记录
							strTempSql = "insert into shop_t (userid, infoId, shopstate, adddate) values (" & lngUserId & "," & thisInfoId & ", 0, '" & now() & "')"
							'response.write strtempsql
							conn.Execute strTempSql
							if err then
							lngstate = 1
							strMsg = "下单失败" 
							end if
						end if
					Next
					
					lngstate = 0
					strMsg = i & "本书下单成功！" & UBound(TempArr)
				else
					lngstate = 1
					strMsg = "会员总借书次数为 0，请联系客服！ 电话：" & gstrServiceTel
				end if
			end if
		end if
	else
		lngstate = 1
		strMsg = "请选中要下单的书本！"
	end if
	if rsFav.state = 1 then rsFav.close
	set rsFav = nothing
	
	Call CloseConn()
%>
{
    "state": "<%=lngstate%>",
    "msg": "<%=strMsg%>"
}