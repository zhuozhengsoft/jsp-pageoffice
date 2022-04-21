<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%@page import="com.zhuozhengsoft.pageoffice.excelwriter.Sheet" %>
<%@page import="com.zhuozhengsoft.pageoffice.excelwriter.Table" %>
<%@page import="com.zhuozhengsoft.pageoffice.excelwriter.Workbook" %>
<%
    //设置PageOfficeCtrl控件的服务页面
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    poCtrl1.setCaption("简单的给Excel赋值");

    //定义Workbook对象
    Workbook workBook = new Workbook();
    //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
    Sheet sheet = workBook.openSheet("Sheet1");
    //定义Table对象，参数“report1”为Excel中定义的名称，“4”为名称指定区域的行数，
    //“5”为名称指定区域的列数，“true”表示表格会按实际数据行数自动扩展
    Table table = sheet.openTableByDefinedName("report", 4, 5, true);
    int rowCount = 12;//假设将要自动填充数据的实际记录条数为12
    for (int i = 1; i <= rowCount; i++) {
        //给区域中的单元格赋值
        table.getDataFields().get(0).setValue(i + "月");
        table.getDataFields().get(1).setValue("100");
        table.getDataFields().get(2).setValue("120");
        table.getDataFields().get(3).setValue("500");
        table.getDataFields().get(4).setValue("120%");
        table.nextRow();//循环下一行，此行必须
    }
    //关闭table对象
    table.close();
    //定义Table对象
    Table table2 = sheet.openTableByDefinedName("report2", 4, 5, true);
    int rowCount2 = 4;//假设将要自动填充数据的实际记录条数为12
    for (int i = 1; i <= rowCount2; i++) {
        //给区域中的单元格赋值
        table2.getDataFields().get(0).setValue(i + "季度");
        table2.getDataFields().get(1).setValue("300");
        table2.getDataFields().get(2).setValue("300");
        table2.getDataFields().get(3).setValue("300");
        table2.getDataFields().get(4).setValue("100%");
        table2.nextRow();
    }
    //关闭table对象
    table2.close();
    poCtrl1.setWriter(workBook);
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    //poCtrl1.setSaveDataPage("SaveData.jsp");
    //poCtrl1.addCustomToolButton("保存", "Save()", 1);
    //打开Word文件
    poCtrl1.webOpen("doc/test4.xls", OpenModeType.xlsNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>给Excel文档中定义名称的区域赋值</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
    <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
        }
    </script>
</head>
<body>
通过openTableByDefinedName填充数据后显示的效果
<div style="width: 1000px; height: 900px;">
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
