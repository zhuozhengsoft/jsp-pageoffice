<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.excelwriter.Sheet"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.excelwriter.Table" %>
<%@ page import="com.zhuozhengsoft.pageoffice.excelwriter.Workbook" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    //定义Workbook对象
    Workbook workBook = new Workbook();
    //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
    Sheet sheet = workBook.openSheet("Sheet1");
    //定义table对象，设置table对象的设置范围
    Table table = sheet.openTable("B4:F13");
    //设置table对象的提交名称，以便保存页面获取提交的数据
    table.setSubmitName("Info");
    poCtrl.setWriter(workBook);
    //添加自定义按钮
    poCtrl.addCustomToolButton("保存", "Save", 1);
    //设置保存页面
    poCtrl.setSaveDataPage("SaveData.jsp");
    //打开Word文档
    poCtrl.webOpen("doc/test.xls", OpenModeType.xlsSubmitForm, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>最简单的提交Excel中的数据</title>
    <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
        }
    </script>
</head>
<body>
<form id="form1">
    <div style="width: auto; height: 700px;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>
