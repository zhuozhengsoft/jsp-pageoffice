<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.excelwriter.Cell" %>
<%@ page import="com.zhuozhengsoft.pageoffice.excelwriter.Sheet" %>
<%@ page import="com.zhuozhengsoft.pageoffice.excelwriter.Workbook" %>
<%@ page import="java.text.DecimalFormat" %>
<%@ page import="java.text.NumberFormat" %>
<%
    //设置PageOfficeCtrl控件的服务页面
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    poCtrl1.setCaption("简单的给Excel赋值");
    //定义Workbook对象
    Workbook workBook = new Workbook();
    //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
    Sheet sheet = workBook.openSheet("Sheet1");
    //定义Cell对象
    Cell cellB4 = sheet.openCell("B4");
    //给单元格赋值
    cellB4.setValue("1月");
    Cell cellC4 = sheet.openCell("C4");
    cellC4.setValue("300");
    Cell cellD4 = sheet.openCell("D4");
    cellD4.setValue("270");
    Cell cellE4 = sheet.openCell("E4");
    cellE4.setValue("270");
    Cell cellF4 = sheet.openCell("F4");
    DecimalFormat df = (DecimalFormat) NumberFormat.getInstance();
    cellF4.setValue(df.format(270.00 / 300 * 100) + "%");

    poCtrl1.setWriter(workBook);
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    //隐藏工具栏
    poCtrl1.setCustomToolbar(false);
    //打开Word文件
    poCtrl1.webOpen("doc/test.xls", OpenModeType.xlsNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>填出Excel表格</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<div style="width: auto; height: 700px;">
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
