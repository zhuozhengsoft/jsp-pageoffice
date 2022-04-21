<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.excelwriter.*,java.awt.*" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    Workbook wb = new Workbook();
    Sheet sheet = wb.openSheet("Sheet1");
    //合并单元格
    sheet.openTable("B2:F2").merge();
    Cell cB2 = sheet.openCell("B2");
    cB2.setValue("北京某公司产品销售情况");
    //设置水平对齐方式
    cB2.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
    cB2.setForeColor(Color.red);
    cB2.getFont().setSize(16);

    sheet.openTable("B4:B6").merge();
    Cell cB4 = sheet.openCell("B4");
    cB4.setValue("A产品");
    //设置水平对齐方式
    cB4.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
    //设置垂直对齐方式
    cB4.setVerticalAlignment(XlVAlign.xlVAlignCenter);

    sheet.openTable("B7:B9").merge();
    Cell cB7 = sheet.openCell("B7");
    cB7.setValue("B产品");
    cB7.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
    cB7.setVerticalAlignment(XlVAlign.xlVAlignCenter);

    poCtrl.setWriter(wb);

    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//此行必须
    poCtrl.setCustomToolbar(false);

    //设置文档打开方式
    poCtrl.webOpen("doc/test.xls", OpenModeType.xlsNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
    <title></title>
    <link href="images/csstg.css" rel="stylesheet" type="text/css"/>
</head>
<body>
<div id="content">
    <div id="textcontent" style="width: 1000px; height: 800px;">
        <!--**************   卓正 PageOffice组件 ************************-->
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</div>
</body>
</html>
