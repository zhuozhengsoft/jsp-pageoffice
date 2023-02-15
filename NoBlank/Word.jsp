<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.WordDocument"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.OfficeVendorType" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    poCtrl1.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>控件与页面底部无空隙</title>
    <style>
        *{
            margin : 0;
            padding : 0;
        }
        div{
            position : absolute;
            width : 100%;
            height : 100%;
        }
        body{
            overflow: hidden;
        }

    </style>

</head>
<body>
<!--**************   卓正 PageOffice组件 ************************-->
<div>
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
