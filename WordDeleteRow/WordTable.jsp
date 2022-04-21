<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType" pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.Cell" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.Table" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>
<%
    String FilePath = request.getContextPath() + "/WordDeleteRow/doc";
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    WordDocument doc = new WordDocument();
    Table table1 = doc.openDataRegion("PO_table").openTable(1);
    Cell cell = table1.openCellRC(2, 1);
    //删除坐标为(2,1)的单元格所在行
    table1.removeRowAt(cell);
    poCtrl.setCustomToolbar(false);
    poCtrl.setWriter(doc);
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.webOpen("doc/test_table.doc", OpenModeType.docNormalEdit, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>删除Word中table中指定单元格所在行</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<div style="color:Red">删除了table中坐标为(2,1)的单元格所在行，请在服务器<%=FilePath%>路径下查看原模板文档。</div>
<div style="width: auto; height: 600px;">
    <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
