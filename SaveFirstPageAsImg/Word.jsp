<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType" pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.PageOfficeCtrl" %>
<%
    // 设置PageOffice组件服务页面
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    poCtrl1.addCustomToolButton("保存", "Save()", 1);
    poCtrl1.addCustomToolButton("保存首页为图片", "SaveFirstAsImg()", 1);
    poCtrl1.setSaveFilePage("SaveFile.jsp");
    // 打开文档
    poCtrl1.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>保存Word首页为图片</title>
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
        if (document.getElementById("PageOfficeCtrl1").CustomSaveResult == "ok") {
            document.getElementById("PageOfficeCtrl1").Alert("文档保存成功!");
        }
    }

    function SaveFirstAsImg() {
        document.getElementById("PageOfficeCtrl1").WebSaveAsImage();
        if (document.getElementById("PageOfficeCtrl1").CustomSaveResult == "ok") {
            document.getElementById("PageOfficeCtrl1").Alert("图片保存成功!");
            document.getElementById("img1").src = "images/test.jpg";
            document.getElementById("img1").style.width = "200px";
            document.getElementById("img1").style.height = "290px";
        }
    }
</script>
<div><img id="img1"/></div>
<div style="width:auto; height:700px;">
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
