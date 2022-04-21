<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType" pageEncoding="utf-8" %>
<%@page import="com.zhuozhengsoft.pageoffice.PageOfficeCtrl" %>
<%
    String userName = request.getParameter("userName");

    if (userName.equals("zhangsan")) userName = "张三";
    if (userName.equals("lisi")) userName = "李四";
    if (userName.equals("wangwu")) userName = "王五";
    //***************************卓正PageOffice组件的使用********************************
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    poCtrl.addCustomToolButton("保存", "Save", 1);
    poCtrl.addCustomToolButton("领导圈阅", "StartHandDraw", 3);
    poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
    poCtrl.setJsFunction_AfterDocumentOpened("ShowByUserName");
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    poCtrl.setSaveFilePage("SaveFile.jsp");
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, userName);
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
            <span style=""> </span><strong>当前用户：</strong>
            <span style="color: Red;"><%=userName %></span>
        </div>
        <script type="text/javascript">
            //保存页面
            function Save() {
                document.getElementById("PageOfficeCtrl1").WebSave();
            }

            //领导圈阅签字
            function StartHandDraw() {
                document.getElementById("PageOfficeCtrl1").HandDraw.Start();
            }

            /*
            //分层显示手写批注
            function ShowHandDrawDispBar() {
                document.getElementById("PageOfficeCtrl1").HandDraw.ShowLayerBar(); ;
            }*/

            //全屏/还原
            function IsFullScreen() {
                document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
            }

            //显示/隐藏用户的手写批注
            function ShowByUserName() {
                var userName = "<%=userName %>"; //从后台获取登录用户名
                document.getElementById("PageOfficeCtrl1").HandDraw.ShowByUserName(null, false); // 隐藏所有的手写
                document.getElementById("PageOfficeCtrl1").HandDraw.ShowByUserName(userName); // 显示Tom的手写，第二个参数为true时可省略
            }
        </script>
        <!--**************   卓正 PageOffice组件 ************************-->
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</div>
</body>
</html>