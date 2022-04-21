<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl, com.zhuozhengsoft.pageoffice.wordwriter.DataRegion"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.DataRegionInsertType" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>
<%
    WordDocument doc = new WordDocument();
    DataRegion mydr1 = doc.createDataRegion("PO_first", DataRegionInsertType.After, "[end]");
    mydr1.selectEnd();
    doc.insertPageBreak();//插入分页符
    DataRegion mydr2 = doc.createDataRegion("PO_second", DataRegionInsertType.After, "[end]");
    mydr2.setValue("[word]doc/test2.doc[/word]");

    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.addCustomToolButton("保存", "Save()", 1);
    poCtrl.setWriter(doc);
    //设置保存页面
    poCtrl.setSaveFilePage("SaveFile.jsp");
    //打开Word文档
    poCtrl.webOpen("doc/test1.doc", OpenModeType.docNormalEdit, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>在word文档中光标处插入分页符</title>
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
        if (document.getElementById("PageOfficeCtrl1").CustomSaveResult == "ok") {
            alert("保存成功！请在/Samples4/InsertPageBreak2/doc目录下查看合并后的新文档\"test3.doc\"。");
        }
    }
</script>
<div style=" width:auto; height:700px;">
    <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
