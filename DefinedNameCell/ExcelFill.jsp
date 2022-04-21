<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%@page import="com.zhuozhengsoft.pageoffice.excelwriter.Sheet" %>
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
    sheet.openCellByDefinedName("testA1").setValue("Tom");
    sheet.openCellByDefinedName("testB1").setValue("John");

    poCtrl1.setWriter(workBook);
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    poCtrl1.setSaveDataPage("SaveData.jsp");
    poCtrl1.addCustomToolButton("保存", "Save()", 1);
    //打开Word文件
    poCtrl1.webOpen("doc/test.xls", OpenModeType.xlsNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>给Excel文档中定义名称的单元格赋值</title>
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
A1、B1单元格的数据是使用后台程序填充进去的，请查看ExcelFill.jsp的代码
<div style="width: 1000px; height: 800px;">
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
