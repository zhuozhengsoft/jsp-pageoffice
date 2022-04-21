<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType" pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.DataRegion" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>
<%
    String userName = request.getParameter("userName");
    //***************************卓正PageOffice组件的使用********************************
    WordDocument doc = new WordDocument();
    //打开数据区域
    DataRegion dTitle = doc.openDataRegion("PO_title");
    //给数据区域赋值
    dTitle.setValue("某公司第二季度产量报表");
    //设置数据区域可编辑性
    dTitle.setEditing(false);//数据区域不可编辑
    DataRegion dA1 = doc.openDataRegion("PO_A_pro1");
    DataRegion dA2 = doc.openDataRegion("PO_A_pro2");
    DataRegion dB1 = doc.openDataRegion("PO_B_pro1");
    DataRegion dB2 = doc.openDataRegion("PO_B_pro2");
    //根据登录用户名设置数据区域可编辑性
    //A部门经理登录后
    if (userName.equals("zhangsan")) {
        userName = "A部门经理";
        dA1.setEditing(true);
        dA2.setEditing(true);
        dB1.setEditing(false);
        dB2.setEditing(false);
    }
    //B部门经理登录后
    else {
        userName = "B部门经理";
        dB1.setEditing(true);
        dB2.setEditing(true);
        dA1.setEditing(false);
        dA2.setEditing(false);
    }
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    poCtrl.setWriter(doc);
    //添加自定义按钮
    poCtrl.addCustomToolButton("保存", "Save", 1);
    poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    //设置保存页
    poCtrl.setSaveFilePage("SaveFile.jsp");
    poCtrl.setMenubar(false);
    //设置文档打开方式
    poCtrl.webOpen("doc/test.doc", OpenModeType.docSubmitForm, userName);
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
            <span style="color: Red;"><%=userName%></span>
        </div>
        <script type="text/javascript">
            //保存页面
            function Save() {
                document.getElementById("PageOfficeCtrl1").WebSave();
            }

            //全屏/还原
            function IsFullScreen() {
                document.getElementById("PageOfficeCtrl1").FullScreen = !document
                    .getElementById("PageOfficeCtrl1").FullScreen;
            }
        </script>
        <!--**************   卓正 PageOffice组件 ************************-->
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</div>

</body>
</html>

