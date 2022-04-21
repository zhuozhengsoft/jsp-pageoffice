<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    //此行必须，设置PageOffice服务器地址
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz");
    //添加自定义按钮
    poCtrl1.addCustomToolButton("保存", "Save", 1);
    poCtrl1.addCustomToolButton("签字", "InsertHandSign()", 2);
    poCtrl1.addCustomToolButton("修改密码", "ChangePsw()", 0);
    //设置保存页面
    poCtrl1.setSaveFilePage("SaveFile.jsp");
    //参数"李志"为当前项目的登录用户名，主要用于痕迹保留。
    poCtrl1.webOpen("test.doc", OpenModeType.docNormalEdit, "李志");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>2.无需输入用户名密码签字</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc; padding: 5px;">
    <span style="color: red;">操作说明：</span>点“签字”按钮即可，插入签字时的用户名为：李志，密码默认为：111111。
    <br/> 关键代码：点右键，选择“查看源文件”，看js函数
    <span style="background-color: Yellow;">InsertHandSign()</span>

</div>
<br/>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }

    function InsertHandSign() {
        try {
            document.getElementById("PageOfficeCtrl1").ZoomSeal.AddHandSign("李志");
        } catch (e) {
        }
    }

    function ChangePsw() {
        document.getElementById("PageOfficeCtrl1").ZoomSeal.ShowSettingsBox();
    }
</script>
<div style="width: auto; height: 700px;">
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>