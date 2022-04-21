<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");
    //打开Word文档
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张佚名");
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>用JS给文档插入图片水印</title>
</head>
<body>
<script type="text/javascript">
    function AfterDocumentOpened() {
        /**
         *document.getElementById("PageOfficeCtrl1").SetWordWaterMarkImage( ImageURL, ImageWidth, IsWashout );
         *ImageURL  字符串类型，必选参数，图片的路径。
         *ImageWidth  整数类型，必选参数，图片的宽度（单位：厘米）。如果是0表示采用图片默认宽度。
         *IsWashout  布尔类型，必选参数，是否冲蚀。true：冲蚀，false：不冲蚀
         */
        document.getElementById("PageOfficeCtrl1").SetWordWaterMarkImage("doc/logo.jpg", 20, false);
    }
</script>
<div style="width:auto; height:670px;">
    <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
