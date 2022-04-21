<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,java.sql.Connection,java.sql.DriverManager,java.sql.ResultSet"
         pageEncoding="utf-8" %>
<%@ page import="java.sql.Statement" %>
<%!
    String subject = "";
    String fileName = "";
%>
<%
    Class.forName("org.sqlite.JDBC");
    String strUrl = "jdbc:sqlite:" + this.getServletContext().getRealPath("demodata/CreateWord.db");
    Connection conn = DriverManager.getConnection(strUrl);
    Statement stmt = conn.createStatement();
    String id = request.getParameter("id");
    if (!id.equals("") && !id.equals(null)) {
        ResultSet rs = stmt.executeQuery("select * from word where ID=" + id);
        subject = rs.getString("Subject");
        fileName = rs.getString("FileName");
        rs.close();
    }
    stmt.close();
    conn.close();
%>
<%
    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    poCtrl1.addCustomToolButton("保存", "Save()", 1);
    //设置保存页面
    poCtrl1.setSaveFilePage("SaveFile.jsp");
    //打开Word文件
    poCtrl1.webOpen("doc/" + fileName, OpenModeType.docNormalEdit, "张三");
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title></title>
    <link href="css/csstg.css" rel="stylesheet" type="text/css"/>
    <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
        }

    </script>
</head>
<body>
<form id="form2">
    <div id="header">
        <div style="float: left; margin-left: 20px;">
            <img src="css/logo.jpg" height="30"/></div>
        <ul>
            <li><a href="#">卓正网站</a></li>
            <li><a href="#">客户问吧</a></li>
            <li><a href="#">在线帮助</a></li>
            <li><a href="#">联系我们</a></li>
        </ul>
    </div>
    <div id="content">
        <div id="textcontent" style="width: 1000px; height: 800px;">
            <div class="flow4">
                <a href="#" onclick="window.external.close();">
                    <img alt="返回" src="css/return.gif" border="0"/>文件列表</a>&nbsp;&nbsp;<strong>文档主题：</strong>
                <span style="color: Red;">
                        <%=subject%>
                        </span> <span style="width: 100px;">
                        </span>
            </div>
            <!-- ****************************PageOffice 组件客户端编程************************************* -->
            <script type="text/javascript">
                function Save() {
                    document.getElementById("PageOfficeCtrl1").WebSave();
                }
            </script>
            <!-- ****************************PageOffice 组件客户端编程结束************************************* -->
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
