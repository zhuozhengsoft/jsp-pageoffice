<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.excelwriter.Sheet"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.excelwriter.Workbook" %>
<%
    Workbook workBook = new Workbook();
    Sheet sheet1 = workBook.openSheet("Sheet1");
    sheet1.openCell("A1").setValue("[image]image/logo.jpg[/image]");
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
//设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.setWriter(workBook);//此行必须
//打开Word文档
    poCtrl.webOpen("doc/test.xls", OpenModeType.xlsNormalEdit, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>Excel中插入图片</title>
</head>
<body>
<form id="form1">
    <div style=" width:100%; height:700px;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>
