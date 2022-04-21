<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType" pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.PageOfficeCtrl, com.zhuozhengsoft.pageoffice.excelwriter.Sheet" %>
<%@ page import="com.zhuozhengsoft.pageoffice.excelwriter.Table" %>
<%@ page import="com.zhuozhengsoft.pageoffice.excelwriter.Workbook" %>

<%
    String userName = request.getParameter("userName");
    //***************************卓正PageOffice组件的使用********************************
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    Workbook wb = new Workbook();
    Sheet sheet = wb.openSheet("Sheet1");
    Table tableA = sheet.openTable("C4:D6");
    Table tableB = sheet.openTable("C7:D9");
    tableA.setSubmitName("tableA");
    tableB.setSubmitName("tableB");
    //根据登录用户名设置数据区域可编辑性
    String strInfo = "";

    //A部门经理登录后
    if (userName.equals("zhangsan")) {
        strInfo = "A部门经理，所以只能编辑A部门的产品数据";
        tableA.setReadOnly(false);
        tableB.setReadOnly(true);
    }
    //B部门经理登录后
    else {
        strInfo = "B部门经理，所以只能编辑B部门的产品数据";
        tableA.setReadOnly(true);
        tableB.setReadOnly(false);
    }
    poCtrl.setWriter(wb);
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    poCtrl.addCustomToolButton("保存", "Save", 1);
    poCtrl.setSaveFilePage("SaveFile.jsp");
    poCtrl.webOpen("doc/test.xls", OpenModeType.xlsSubmitForm, userName);
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
        <div class="flow4">
            <a href="Default.jsp"> 返回登录页</a>
            <strong>当前用户：</strong>
            <span style="color: Red;"><%=strInfo %></span>
        </div>

        <script type="text/javascript">
            //保存页面
            function Save() {
                document.getElementById("PageOfficeCtrl1").WebSave();
            }

        </script>
        <!--**************   卓正 PageOffice组件 ************************-->
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</div>
</body>
</html>