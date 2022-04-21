<%@ page language="java" pageEncoding="utf-8" %>
<%
    String path = request.getContextPath();
    String basePath = request.getScheme() + "://" + request.getServerName() + ":" + request.getServerPort() + path + "/";
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>登录页面</title>
</head>
<body>
<form id="form1" action="SetHandDrawByUserName.jsp" method="post">
    <div style=" text-align:center;">
        <div>请选择登录用户：</div>
        <br/>
        <select name="userName">
            <option selected="selected" value="zhangsan">张三</option>
            <option value="lisi">李四</option>
            <option value="wangwu">王五</option>
        </select><br/><br/>
        <input type="submit" value="打开文件"/><br/><br/>
        <div style=" color:Red;">不同的用户登录后，在文档中隐藏其他人的手写批注，只显示当前用户的手写批注。</div>
    </div>
</form>
</body>
</html>
