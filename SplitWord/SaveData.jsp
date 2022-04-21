<%@ page language="java" import="com.zhuozhengsoft.pageoffice.wordreader.DataRegion,com.zhuozhengsoft.pageoffice.wordreader.WordDocument,java.io.FileOutputStream"
         pageEncoding="utf-8" %>
<%
    String filePath = request.getSession().getServletContext().getRealPath("SplitWord/doc/") + "/";
    WordDocument doc = new WordDocument(request, response);
    byte[] bWord;

    DataRegion dr1 = doc.openDataRegion("PO_test1");
    bWord = dr1.getFileBytes();
    FileOutputStream fos1 = new FileOutputStream(filePath + "new1.doc");
    fos1.write(bWord);
    fos1.flush();
    fos1.close();

    DataRegion dr2 = doc.openDataRegion("PO_test2");
    bWord = dr2.getFileBytes();
    FileOutputStream fos2 = new FileOutputStream(filePath + "new2.doc");
    fos2.write(bWord);
    fos2.flush();
    fos2.close();

    DataRegion dr3 = doc.openDataRegion("PO_test3");
    bWord = dr3.getFileBytes();
    FileOutputStream fos3 = new FileOutputStream(filePath + "new3.doc");
    fos3.write(bWord);
    fos3.flush();
    fos3.close();

    //doc.showPage(500,400);
    doc.close();
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>My JSP 'SaveFile.jsp' starting page</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
    <!--
    <link rel="stylesheet" type="text/css" href="styles.css">
    -->
</head>
<body>
</body>
</html>
