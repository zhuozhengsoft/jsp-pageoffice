<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,javax.servlet.*"
         pageEncoding="utf-8" %>
<%
    String fileName = "test.doc"; // 正文内容文件：test.doc

    //***************************卓正PageOffice组件的使用********************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.addCustomToolButton("保存", "Save", 1);
    poCtrl1.addCustomToolButton("领导圈阅", "StartHandDraw", 3);
    poCtrl1.addCustomToolButton("分层显示手写批注", "ShowHandDrawDispBar", 7);
    poCtrl1.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz");//此行必须
    poCtrl1.setSaveFilePage("savefile.jsp?fileName=test.doc");
    poCtrl1.webOpen("doc/" + fileName, OpenModeType.docNormalEdit, "张三");
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
    <title></title>
    <link href="images/csstg.css" rel="stylesheet" type="text/css"/>
</head>
<body>
<form id="form2">
    <div id="header">
        <div style="float: left; margin-left: 20px;">
            <img src="images/logo.jpg" height="30"/></div>
        <ul>
            <li><a target="_blank" href="http://www.zhuozhengsoft.com">卓正网站</a></li>
            <li><a target="_blank" href="http://www.zhuozhengsoft.com/poask/index.asp">客户问吧</a></li>
            <li><a href="#">在线帮助</a></li>
            <li><a target="_blank" href="http://www.zhuozhengsoft.com/about/about/">联系我们</a></li>
        </ul>
    </div>
    <div id="content">
        <div id="textcontent" style="width: 1000px; height: 800px;">
            <div class="flow4">
                <a href="#" onclick="window.external.close();"> <img alt="返回" src="images/return.gif"
                                                                     border="0"/>文件列表</a>
                <span style="width: 100px;"> </span><strong>文档主题：</strong>
                <span style="color: Red;">测试文件</span>
            </div>
            <script type="text/javascript">
                function Save() {
                    document.getElementById("PageOfficeCtrl1").WebSave();
                }

                //领导圈阅签字
                function StartHandDraw() {
                    document.getElementById("PageOfficeCtrl1").HandDraw.SetPenWidth(5);
                    document.getElementById("PageOfficeCtrl1").HandDraw.Start();
                }

                //分层显示手写批注
                function ShowHandDrawDispBar() {
                    document.getElementById("PageOfficeCtrl1").HandDraw.ShowLayerBar();
                    ;
                }

                //全屏/还原
                function IsFullScreen() {
                    document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
                }

            </script>
            <!--**************   卓正 PageOffice组件 ************************-->
            <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
        </div>
    </div>
    <div id="footer">
        <hr width="1000"/>
        <div>
            Copyright (c) 2019 北京卓正志远软件有限公司
        </div>
    </div>
</form>
</body>
</html>
