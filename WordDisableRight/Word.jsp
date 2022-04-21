<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.WordDocument,javax.servlet.*,javax.servlet.http.*"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    WordDocument wordDoc = new WordDocument();
    wordDoc.setDisableWindowRightClick(true);//禁止word鼠标右键
    //wordDoc.setDisableWindowDoubleClick(true);//禁止word鼠标双击
    //wordDoc.setDisableWindowSelection(true);//禁止在word中选择文件内容
    poCtrl1.setWriter(wordDoc);
    //打开文件
    poCtrl1.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>禁止Word文档中鼠标右键</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<div style="color:Red">打开Word文档后，鼠标右键，发现右键失效。</div>
<%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</body>
</html>
