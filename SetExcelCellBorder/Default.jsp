<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.excelwriter.*,java.awt.*" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    Workbook wb = new Workbook();
    Sheet sheet = wb.openSheet("Sheet1");
    // 设置背景
    Table backGroundTable = sheet.openTable("A1:P200");
    //设置表格边框样式
    backGroundTable.getBorder().setLineColor(Color.white);

    // 设置单元格边框样式
    Border C4Border = sheet.openTable("C4:C4").getBorder();
    C4Border.setWeight(XlBorderWeight.xlThick);
    C4Border.setLineColor(Color.yellow);
    C4Border.setBorderType(XlBorderType.xlAllEdges);

    // 设置单元格边框样式
    Border B6Border = sheet.openTable("B6:B6").getBorder();
    B6Border.setWeight(XlBorderWeight.xlHairline);
    B6Border.setLineColor(Color.magenta);
    B6Border.setLineStyle(XlBorderLineStyle.xlSlantDashDot);
    B6Border.setBorderType(XlBorderType.xlAllEdges);

    //设置表格边框样式
    Table titleTable = sheet.openTable("B4:F5");
    titleTable.getBorder().setWeight(XlBorderWeight.xlThick);
    titleTable.getBorder().setLineColor(new Color(0, 128, 128));
    titleTable.getBorder().setBorderType(XlBorderType.xlAllEdges);

    //设置表格边框样式
    Table bodyTable2 = sheet.openTable("B6:F15");
    bodyTable2.getBorder().setWeight(XlBorderWeight.xlThick);
    bodyTable2.getBorder().setLineColor(new Color(0, 128, 128));
    bodyTable2.getBorder().setBorderType(XlBorderType.xlAllEdges);

    poCtrl.setWriter(wb);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
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
