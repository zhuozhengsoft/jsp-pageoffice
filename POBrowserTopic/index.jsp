<%@ page language="java" pageEncoding="utf-8" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head id="Head1">
    <title></title>
    <link href="css/style.css" rel="stylesheet" type="text/css"/>
    <script type="text/javascript" src="../js/jquery-1.8.2.min.js"></script>
    <!--pageoffice.js文件一定要引用-->
    <script type="text/javascript" src="../pageoffice.js"></script>
    <script type="text/javascript">
        function getText() {
            //获取文本框的值，传递给result2.jsp
            var iptValue = document.getElementById("ipt").value;
            $.ajax({
                type: "POST",  //提交方式
                url: "Result2.jsp?param=" + iptValue,//路径
                dataType: "text",
                success: function (data) {//返回数据根据结果进行相应的处理
                }
            });
        }
    </script>
</head>
<%
    String userName = "张三";
    session.setAttribute("userName", userName);
%>
<body>
<!--header-->
<div class="zz-headBox br-5 clearfix" align="center">
    <div class="zz-head mc">
        <!--logo-->
        <div class="logo fl">
            <a href="#">
                <img src="images/logo.png" alt=""/></a></div>
        <!--logo end-->
        <ul class="head-rightUl fr">
            <li><a href="http://www.zhuozhengsoft.com">卓正网站</a></li>
            <li><a href="http://www.zhuozhengsoft.com/Technical/">技术支持</a></li>
            <li class="bor-0"><a href="http://www.zhuozhengsoft.com/about/about/">联系我们</a></li>
        </ul>
    </div>
</div>
<!--header end-->
<!--content-->
<div class="zz-content mc clearfix pd-28" align="center">
    <div class="demo mc">
        <h2 class="fs-16">
            POBrowser 专题</h2>
    </div>

    <div style="margin : 10px" align="center">
        <table style="border-collapse:separate; border-spacing:20px; ">
            <tr>
                <td>1.POBrowser方式最简单的打开编辑保存文档：<a
                        href="javascript:POBrowser.openWindowModeless('Word1.jsp' , 'width=1200px;height=800px;');">Word1.jsp</a>
                </td>
            </tr>
            <tr>
                <td>2.向POBrowser打开的页面传递参数1：<a
                        href="javascript:POBrowser.openWindowModeless('Word2.jsp?id=1' , 'width=1200px;height=800px;');">Word2.jsp</a>
                </td>
            </tr>
			<!--
            <tr>
                <td>3.向POBrowser打开的页面传递参数2：<a
                        href="javascript:POBrowser.openWindowModeless('Word3.jsp' , 'width=1200px;height=800px;');"
                        onclick="getText()">Word3.jsp</a></td>
            <tr>
                <td>请输入值：
                    <input type="text" id="ipt" value="1234"/>
                <td>
            </tr>
			-->
        </table>
    </div>
</div>
</body>
</html>
