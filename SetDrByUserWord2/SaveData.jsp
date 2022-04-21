<%@ page language="java" import="com.zhuozhengsoft.pageoffice.wordreader.WordDocument" pageEncoding="utf-8" %>
<%@page import="java.io.FileOutputStream" %>
<%
    //-----------  PageOffice 服务器端编程开始  -------------------//
    WordDocument doc = new WordDocument(request, response);
    byte[] bytes = null;
    String filePath = "";
    if (request.getParameter("userName") != null && request.getParameter("userName").trim().equalsIgnoreCase("zhangsan")) {
        bytes = doc.openDataRegion("PO_com1").getFileBytes();
        filePath = "content1.doc";
    } else {
        bytes = doc.openDataRegion("PO_com2").getFileBytes();
        filePath = "content2.doc";
    }
    doc.close();

    filePath = request.getSession().getServletContext().getRealPath("SetDrByUserWord2/doc/") + "/" + filePath;
    FileOutputStream outputStream = new FileOutputStream(filePath);
    outputStream.write(bytes);
    outputStream.flush();
    outputStream.close();
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>登录页面</title>
</head>
<body>
</body>
</html>
