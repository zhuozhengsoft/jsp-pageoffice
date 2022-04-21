<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    //设置PageOffice控件标题栏内容
    poCtrl.setCaption("PageOfficeCtrl控件标题栏内容");
    //打开Word文档
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>演示：修改PageOffice控件标题栏文本内容</title>
    <style>
        html, body {
            height: 100%;
        }

        .main {
            height: 100%;
        }
    </style>
</head>
<body>
<form id="form1">
    <div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc;
        padding: 5px;">
        操作：在Html页面添加PageOfficeCtrl控件，再在后台设置PageOfficeCtrl对象的Caption属性<br/>
        关键代码：<span style="background-color:Yellow;">poCtrl.setCaption("PageOfficeCtrl控件的标题栏内容");//poCtrl为PageOfficeCtrl对象</span>
    </div>
    <div style="height: 600px; width: auto;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>