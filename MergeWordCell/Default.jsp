<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.DataRegion,com.zhuozhengsoft.pageoffice.wordwriter.Table" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WdParagraphAlignment" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>
<%@ page import="java.awt.*" %>
<%
    WordDocument doc = new WordDocument();
    DataRegion dataReg = doc.openDataRegion("PO_table");
    Table table = dataReg.openTable(1);
    //合并table中的单元格
    table.openCellRC(1, 1).mergeTo(1, 4);
    //给合并后的单元格赋值
    table.openCellRC(1, 1).setValue("销售情况表");
    //设置单元格文本样式
    table.openCellRC(1, 1).getFont().setColor(Color.red);
    table.openCellRC(1, 1).getFont().setSize(24);
    table.openCellRC(1, 1).getFont().setName("楷体");
    table.openCellRC(1, 1).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    poCtrl.setWriter(doc);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//此行必须
    poCtrl.setCustomToolbar(false);
    //设置文档打开方式
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张三");
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