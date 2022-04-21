<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    //添加自定义按钮
    poCtrl.addCustomToolButton("保存", "Save", 1);
    poCtrl.addCustomToolButton("打印设置", "PrintSet", 0);
    poCtrl.addCustomToolButton("打印", "PrintFile", 6);
    poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
    poCtrl.addCustomToolButton("-", "", 0);
    poCtrl.addCustomToolButton("关闭", "Close", 21);
    //设置保存页面
    poCtrl.setSaveFilePage("SaveFile.jsp");
    //打开Word文档
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张佚名");
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
    <meta charset="utf-8">
    <title>XX文档系统</title>
    <style>
        #main {
            width: 768px;
            height: 887px;
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
<body style="overflow:hidden">
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }

    function Close() {
        window.external.close();
    }
</script>
<form id="form1">
    <div id="main">
        <div id="shut"><img src="../images/close.png" onclick="Close()" title="关闭"/></div>
        <div id="content" style="height:850px;width:768px;overflow-y:hidden;">
            <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
        </div>
    </div>
</form>
</body>
</html>
