<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.DataRegion"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.addCustomToolButton("保存图片", "Save", 1);
    WordDocument wordDoc = new WordDocument();
    //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
    DataRegion dataRegion1 = wordDoc.openDataRegion("PO_image");
    dataRegion1.setEditing(true);//放图片的数据区域是可以编辑的，其它部分不可编辑
    poCtrl.setWriter(wordDoc);
    //设置保存页面
    poCtrl.setSaveDataPage("SaveData.jsp");
    poCtrl.webOpen("doc/test.doc", OpenModeType.docSubmitForm, "张佚名");
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>保存时获取word文档中的图片</title>
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
        var value = document.getElementById("PageOfficeCtrl1").CustomSaveResult;
        document.getElementById("PageOfficeCtrl1").Alert(value);
    }
</script>
<div id="div1" style="width: auto; height: 700px;">
    <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
