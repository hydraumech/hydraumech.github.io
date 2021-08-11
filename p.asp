<%
Dim url, body, myCache

url = Request.QueryString("url")

  Set myCache = new cache
  myCache.name = "picindex"&url
  If myCache.valid Then
          body = myCache.value
  Else
          body = GetWebData(url)
          myCache.add body,dateadd("d",1,now)
  End If

  If Err.Number = 0 Then
        Response.CharSet = "UTF-8"
        Response.ContentType = "application/octet-stream"
        Response.BinaryWrite body
        Response.Flush
  Else
        Wscript.Echo Err.Description
  End if

'取得数据
Public Function GetWebData(ByVal strUrl)
Dim curlpath
curlpath = Mid(strUrl,1,Instr(8,strUrl,"/"))
Dim Retrieval
Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
With Retrieval
.Open "Get", strUrl, False,"",""
.setRequestHeader "Referer", curlpath
.Send
GetWebData =.ResponseBody
End With
Set Retrieval = Nothing
End Function

'cache类

class Cache
        private obj                                'cache内容
        private expireTime                '过期时间
        private expireTimeName        '过期时间application名
        private cacheName                'cache内容application名
        private path                        'url
        
        private sub class_initialize()
                path=request.servervariables("url")
                path=left(path,instrRev(path,"/"))
        end sub
        
        private sub class_terminate()
        end sub
        
        public property get blEmpty
                '是否为空
                if isempty(obj) then
                        blEmpty=true
                else
                        blEmpty=false
                end if
        end property
        
        public property get valid
                '是否可用(过期)
                if isempty(obj) or not isDate(expireTime) then
                        valid=false
                elseif CDate(expireTime)<now then
                                valid=false
                else
                        valid=true
                end if
        end property
        
        public property let name(str)
                '设置cache名
                cacheName=str & path
                obj=application(cacheName)
                expireTimeName=str & "expires" & path
                expireTime=application(expireTimeName)
        end property
        
        public property let expires(tm)
                '重设置过期时间
                expireTime=tm
                application.lock
                application(expireTimeName)=expireTime
                application.unlock
        end property
        
        public sub add(var,expire)
                '赋值
                if isempty(var) or not isDate(expire) then
                        exit sub
                end if
                obj=var
                expireTime=expire
                application.lock
                application(cacheName)=obj
                application(expireTimeName)=expireTime
                application.unlock
        end sub
        
        public property get value
                '取值
                if isempty(obj) or not isDate(expireTime) then
                        value=null
                elseif CDate(expireTime)<now then
                        value=null
                else
                        value=obj
                end if
        end property
        
        public sub makeEmpty()
                '释放application
                application.lock
                application(cacheName)=empty
                application(expireTimeName)=empty
                application.unlock
                obj=empty
                expireTime=empty
        end sub
        
        public function equal(var2)
                '比较
                if typename(obj)<>typename(var2) then
                        equal=false
                elseif typename(obj)="Object" then
                        if obj is var2 then
                                equal=true
                        else
                                equal=false
                        end if
                elseif typename(obj)="Variant()" then
                        if join(obj,"^")=join(var2,"^") then
                                equal=true
                        else
                                equal=false
                        end if
                else
                        if obj=var2 then
                                equal=true
                        else
                                equal=false
                        end if
                end if
        end function
end class
%>