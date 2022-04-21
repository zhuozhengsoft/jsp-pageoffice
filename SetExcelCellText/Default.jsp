<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType, com.zhuozhengsoft.pageoffice.PageOfficeCtrl" pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.excelwriter.*,java.awt.*" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    Workbook wb = new Workbook();
    Sheet sheet = wb.openSheet("Sheet1");

    Cell cC3 = sheet.openCell("C3");
    //设置单元格背景样式
    cC3.setBackColor(Color.LIGHT_GRAY);
    cC3.setValue("一月");
    cC3.setForeColor(Color.white);
    cC3.setHorizontalAlignment(XlHAlign.xlHAlignCenter);

    Cell cD3 = sheet.openCell("D3");
    //设置单元格背景样式
    cD3.setBackColor(Color.lightGray);
    cD3.setValue("二月");
    cD3.setForeColor(Color.white);
    cD3.setHorizontalAlignment(XlHAlign.xlHAlignCenter);

    Cell cE3 = sheet.openCell("E3");
    //设置单元格背景样式
    cE3.setBackColor(Color.lightGray);
    cE3.setValue("三月");
    cE3.setForeColor(Color.white);
    cE3.setHorizontalAlignment(XlHAlign.xlHAlignCenter);

    Cell cB4 = sheet.openCell("B4");
    //设置单元格背景样式
    cB4.setBackColor(new Color(10, 254, 254));
    cB4.setValue("住房");
    cB4.setForeColor(new Color(10, 150, 150));
    cB4.setHorizontalAlignment(XlHAlign.xlHAlignCenter);

    Cell cB5 = sheet.openCell("B5");
    //设置单元格背景样式
    cB5.setBackColor(new Color(10, 150, 150));
    cB5.setValue("三餐");
    cB5.setForeColor(new Color(10, 100, 250));
    cB5.setHorizontalAlignment(XlHAlign.xlHAlignCenter);

    Cell cB6 = sheet.openCell("B6");
    //设置单元格背景样式
    cB6.setBackColor(new Color(200, 200, 100));
    cB6.setValue("车费");
    cB6.setForeColor(new Color(10, 150, 150));
    cB6.setHorizontalAlignment(XlHAlign.xlHAlignCenter);

    Cell cB7 = sheet.openCell("B7");
    //设置单元格背景样式
    cB7.setBackColor(new Color(80, 50, 80));
    cB7.setValue("通讯");
    cB7.setForeColor(new Color(10, 150, 150));
    cB7.setHorizontalAlignment(XlHAlign.xlHAlignCenter);

    //绘制表格线
    Table titleTable = sheet.openTable("B3:E10");
    titleTable.getBorder().setWeight(XlBorderWeight.xlThick);
    titleTable.getBorder().setLineColor(new Color(0, 128, 128));
    titleTable.getBorder().setBorderType(XlBorderType.xlAllEdges);

    sheet.openTable("B1:E2").merge();//合并单元格
    sheet.openTable("B1:E2").setRowHeight(30);//设置行高
    Cell B1 = sheet.openCell("B1");
    //设置单元格文本样式
    B1.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
    B1.setVerticalAlignment(XlVAlign.xlVAlignCenter);
    B1.setForeColor(new Color(0, 128, 128));
    B1.setValue("出差开支预算");
    B1.getFont().setBold(true);
    B1.getFont().setSize(25);

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
