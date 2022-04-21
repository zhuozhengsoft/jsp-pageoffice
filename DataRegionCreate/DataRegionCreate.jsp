<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.DataRegion"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.DataRegionInsertType" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>
<%
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须

    WordDocument doc = new WordDocument();
    //创建数据区域，createDataRegion 方法中的三个参数分别代表“新建的数据区域名称”，“数据区域将要插入的位置”，
    //“与新建的数据区域相关联的数据区域名称”，若当前Word文档中尚无数据区域（书签）或者想在文档的最开头创建时，那么第三个参数为“[home]”
    //若想在文档的结尾处创建数据区域则第三个参数为“[end]”
    DataRegion dataRegion1 = doc.createDataRegion("reg1", DataRegionInsertType.After, "[home]");
    //设置创建的数据区域的可编辑性
    dataRegion1.setEditing(true);
    //给数据区域赋值
    dataRegion1.setValue("第一个数据区域\r\n");
    DataRegion dataRegion2 = doc.createDataRegion("reg2", DataRegionInsertType.After, "reg1");
    dataRegion2.setEditing(true);
    dataRegion2.setValue("第二个数据区域");

    poCtrl1.setWriter(doc);
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    //隐藏工具栏
    poCtrl1.setCustomToolbar(false);
    //打开Word文件
    poCtrl1.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>新建数据区域</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">

</head>
<body>
<div style="width: auto; height: 700px;">
    <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
</div>
</body>
</html>
