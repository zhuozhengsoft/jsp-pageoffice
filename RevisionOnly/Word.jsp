<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    //添加自定义按钮
    poCtrl.addCustomToolButton("保存", "Save", 1);
    poCtrl.addCustomToolButton("隐藏痕迹", "hideRevision", 18);
    poCtrl.addCustomToolButton("显示痕迹", "showRevision", 9);
    //设置保存页面
    poCtrl.setSaveFilePage("SaveFile.jsp");
    //打开Word文档
    poCtrl.webOpen("doc/test.doc", OpenModeType.docRevisionOnly, "李斯");
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
    <meta charset="utf-8">
    <title>XX文档系统</title>
    <style>
        #main {
            width: 1040px;
            height: 890px;
            border: #83b3d9 2px solid;
            background: #f2f7fb;

        }
        #shut {
            width: 45px;
            height: 30px;
            float: right;
            margin-right: -1px;
        }
        #shut:hover {
        }
    </style>
</head>
<body style="margin:0; padding:0;border:0px; overflow:hidden" scroll="no">
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }

    function showRevision() {
        document.getElementById("PageOfficeCtrl1").ShowRevisions = true;
    }

    function hideRevision() {
        document.getElementById("PageOfficeCtrl1").ShowRevisions = false;
    }
</script>
<div id="main">
    <div id="content" style="height:850px;width:1036px;overflow:hidden;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</div>
</body>
</html>

