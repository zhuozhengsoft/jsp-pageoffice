<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.excelwriter.Sheet"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.excelwriter.Table" %>
<%@ page import="com.zhuozhengsoft.pageoffice.excelwriter.Workbook" %>
<%
    //设置PageOfficeCtrl控件的服务页面
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    poCtrl1.setCaption("使用OpenTable给Excel赋值");
    //定义Workbook对象
    Workbook workBook = new Workbook();
    //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
    Sheet sheet = workBook.openSheet("Sheet1");
    //定义Table对象
    Table table = sheet.openTable("B4:F13");
    for (int i = 0; i < 50; i++) {
        table.getDataFields().get(0).setValue("产品 " + i);
        table.getDataFields().get(1).setValue("100");
        table.getDataFields().get(2).setValue(String.valueOf(100 + i));
        table.nextRow();
    }
    table.close();

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
    <title></title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
此表格中的数据是使用后台程序填充进去的，请查看Table.jsp的代码，使用的OpenTable的方法，可以实现行增长，还可以循环使用原模板Table区域（B4:F13）单元格样式。
<div style="width: 1000px; height: 700px;">
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
