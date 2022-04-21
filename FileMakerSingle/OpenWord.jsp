<%@ page language="java" import="com.zhuozhengsoft.pageoffice.*" pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setCustomToolbar(false);
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.webOpen("doc/maker.doc", OpenModeType.docReadOnly, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
</head>
<body style="overflow:hidden">
<!--**************   卓正 PageOffice 客户端代码开始    ************************-->
<div style="height:850px;width:auto;">
    <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>