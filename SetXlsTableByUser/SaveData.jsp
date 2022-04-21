<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.excelreader.Sheet,com.zhuozhengsoft.pageoffice.excelreader.Table,com.zhuozhengsoft.pageoffice.excelreader.Workbook"
         pageEncoding="utf-8" %>
<%
    Workbook wb = new Workbook(request, response);
    Sheet sheet = wb.openSheet("Sheet1");
    Table tableA = sheet.openTable("tableA");
    Table tableB = sheet.openTable("tableB");

    //输出提交的数据
    out.print("提交的数据为：<br/><br/>");
    out.print("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;计划完成量"
            + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;实际完成量<br/>");
    out.print("A部门：");
    while (!tableA.getEOF()) {
        if (!tableA.getDataFields().getIsEmpty()) {
            for (int i = 0; i < tableA.getDataFields().size(); i++) {
                out.print("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        + tableA.getDataFields().get(i).getValue());
            }
        }
        out.print("<br/>");
        tableA.nextRow();
    }

    out.print("B部门：");
    while (!tableB.getEOF()) {
        if (!tableB.getDataFields().getIsEmpty()) {
            for (int i = 0; i < tableB.getDataFields().size(); i++) {
                out.print("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        + tableB.getDataFields().get(i).getValue());
            }
        }
        out.print("<br/>");
        tableB.nextRow();
    }

    //向客户端显示提交的数据
    wb.showPage(300, 300);
    wb.close();
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
