<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType" pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.*" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    poCtrl.setCustomToolbar(false);//隐藏用户自定义工具栏
    WordDocument doc = new WordDocument();
    //在word中指定的"PO_table1"的数据区域内动态创建一个3行5列的表格
    Table table1 = doc.openDataRegion("PO_table1").createTable(3, 5, WdAutoFitBehavior.wdAutoFitWindow);
    //合并(1,1)到(3,1)的单元格并赋值
    table1.openCellRC(1, 1).mergeTo(3, 1);
    table1.openCellRC(1, 1).setValue("合并后的单元格");
    //给表格table1中剩余的单元格赋值
    for (int i = 1; i < 4; i++) {
        table1.openCellRC(i, 2).setValue("AA" + String.valueOf(i));
        table1.openCellRC(i, 3).setValue("BB" + String.valueOf(i));
        table1.openCellRC(i, 4).setValue("CC" + String.valueOf(i));
        table1.openCellRC(i, 5).setValue("DD" + String.valueOf(i));
    }

    //在"PO_table1"后面动态创建一个新的数据区域"PO_table2",用于创建新的一个5行5列的表格table2
    DataRegion drTable2 = doc.createDataRegion("PO_table2", DataRegionInsertType.After, "PO_table1");
    Table table2 = drTable2.createTable(5, 5, WdAutoFitBehavior.wdAutoFitWindow);
    //给新表格table2赋值
    for (int i = 1; i < 6; i++) {
        table2.openCellRC(i, 1).setValue("AA" + String.valueOf(i));
        table2.openCellRC(i, 2).setValue("BB" + String.valueOf(i));
        table2.openCellRC(i, 3).setValue("CC" + String.valueOf(i));
        table2.openCellRC(i, 4).setValue("DD" + String.valueOf(i));
        table2.openCellRC(i, 5).setValue("EE" + String.valueOf(i));
    }
    poCtrl.setWriter(doc);//此行必须
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.webOpen("doc/createTable.doc", OpenModeType.docNormalEdit, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>Word中动态创建表格</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<div style="width: auto; height: 800px;">
    <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
