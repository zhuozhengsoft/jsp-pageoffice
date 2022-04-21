<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="UTF-8" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    //此行必须，设置PageOffice服务器地址
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz");
    //添加自定义按钮
    poCtrl1.addCustomToolButton("保存", "Save", 1);
    poCtrl1.addCustomToolButton("签字", "AddHandSign()", 3);
    poCtrl1.addCustomToolButton("验证印章", "VerifySeal()", 5);
    //设置保存页面
    poCtrl1.setSaveFilePage("SaveFile.jsp");
    poCtrl1.webOpen("test.xls", OpenModeType.xlsNormalEdit, "李志");//参数"李志"为您开发项目的登录用户名。
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>3.特殊签字需求实现：签批后签字不保护文档内容，用户仍可编辑修改，签字不会出现失效字样。</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc; padding: 5px;">
    <span style="color: red;">操作说明：</span>点“签字”按钮即可，插入印章时的用户名为：李志，密码默认为：111111。
    <br/> 关键代码：点右键，选择“查看源文件”，看js函数
    <span style="background-color: Yellow;"> AddHandSign()</span>
</div>
<br/>
<br/>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }

    function AddHandSign() {
        try {
            document.getElementById("PageOfficeCtrl1").ZoomSeal.AddHandSign("", "");//第二个参数为空字符串时，签完名的文档是不受保护的，即签完名后还可以进行编辑等操作；第二个参数为null时，签完名的文档是受保护的，即签完名后不能再进行编辑等操作。
        } catch (e) {
        }
        ;
    }

    function VerifySeal() {
        document.getElementById("PageOfficeCtrl1").ZoomSeal.VerifySeal();
    }
</script>
<div style="width: auto; height: 700px;">
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>