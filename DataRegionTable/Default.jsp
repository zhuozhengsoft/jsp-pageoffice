<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType, com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%@page import="com.zhuozhengsoft.pageoffice.wordwriter.*, java.awt.*" %>
<%
    //***************************卓正PageOffice组件的使用********************************
    WordDocument doc = new WordDocument();
    //打开数据区域
    DataRegion dTable = doc.openDataRegion("PO_table");
    //设置数据区域可编辑性
    dTable.setEditing(true);
    //打开数据区域中的表格，OpenTable(index)方法中的index为word文档中表格的下标，从1开始
    Table table1 = doc.openDataRegion("PO_Table").openTable(1);
    //设置表格边框样式
    table1.getBorder().setLineColor(Color.green);
    table1.getBorder().setLineWidth(WdLineWidth.wdLineWidth050pt);
    // 设置表头单元格文本居中
    table1.openCellRC(1, 2).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
    table1.openCellRC(1, 3).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
    table1.openCellRC(2, 1).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
    table1.openCellRC(3, 1).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
    // 给表头单元格赋值
    table1.openCellRC(1, 2).setValue("产品1");
    table1.openCellRC(1, 3).setValue("产品2");
    table1.openCellRC(2, 1).setValue("A部门");
    table1.openCellRC(3, 1).setValue("B部门");

    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    poCtrl.setWriter(doc);
    //添加自定义按钮
    poCtrl.addCustomToolButton("保存", "Save", 1);
    poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    //设置保存页
    poCtrl.setSaveDataPage("SaveData.jsp");
    //设置文档打开方式
    poCtrl.webOpen("doc/test.doc", OpenModeType.docSubmitForm, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
    <title>数据区域提交表格</title>
    <link href="images/csstg.css" rel="stylesheet" type="text/css"/>
</head>
<body>
<div id="content">
    <div id="textcontent" style="width: 1000px; height: 800px;">
        <script type="text/javascript">
            //保存
            function Save() {
                document.getElementById("PageOfficeCtrl1").WebSave();
            }

            //全屏/还原
            function IsFullScreen() {
                document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
            }
        </script>
        <!--**************   卓正 PageOffice组件 ************************-->
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</div>
</body>
</html>
