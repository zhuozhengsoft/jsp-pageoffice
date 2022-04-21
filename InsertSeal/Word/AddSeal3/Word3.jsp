<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    //添加自定义按钮
    poCtrl1.addCustomToolButton("保存", "Save", 1);
    poCtrl1.addCustomToolButton("加盖印章", "InsertSeal()", 2);
    //设置保存页面
    poCtrl1.setSaveFilePage("SaveFile.jsp");
    poCtrl1.webOpen("word3.doc", OpenModeType.docNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>3.不弹出用户名、密码输入框并且不弹出印章选择框盖章</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }

    function InsertSeal() {
        try {
            /**
             *第一个参数，必填项，标识印章名称（当存在重名的印章时，默认选取第一个印章）；
             *第二个参数，可选项，标识是否保护文档，为null时保护文档，为空字符串时不保护文档；
             *第三个参数，可选项，标识盖章指定位置名称，须为英文或数字，不区分大小写
             */
            var bRet = document.getElementById("PageOfficeCtrl1").ZoomSeal.AddSealByName("测试公章", null, null);
            if (bRet) {
                alert("盖章成功！");
            } else {
                alert("盖章失败！");
            }
        } catch (e) {
        }
    }
</script>
<div style="width: auto; height: 700px;">
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>

</html>