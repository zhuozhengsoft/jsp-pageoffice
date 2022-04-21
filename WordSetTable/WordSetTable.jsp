<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.DataRegion"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.Table" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须

    WordDocument doc = new WordDocument();
    //打开数据区域
    DataRegion dataRegion = doc.openDataRegion("PO_regTable");
    //打开table，openTable(index)方法中的index代表Word文档中table位置的索引，从1开始
    Table table = dataRegion.openTable(1);
    //给table中的单元格赋值， openCellRC(int,int)中的参数分别代表第几行、第几列，从1开始
    table.openCellRC(3, 1).setValue("A公司");
    table.openCellRC(3, 2).setValue("开发部");
    table.openCellRC(3, 3).setValue("李清");
    //插入一行，insertRowAfter方法中的参数代表在哪个单元格下面插入一个空行
    table.insertRowAfter(table.openCellRC(3, 3));
    table.openCellRC(4, 1).setValue("B公司");
    table.openCellRC(4, 2).setValue("销售部");
    table.openCellRC(4, 3).setValue("张三");

    poCtrl1.setWriter(doc);
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    //隐藏自定义工具栏
    poCtrl1.setCustomToolbar(false);
    //打开文件
    poCtrl1.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>简单的给Word文档中的Table赋值</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<div style=" width:auto; height:700px;">
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
