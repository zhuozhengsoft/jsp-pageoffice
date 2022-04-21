<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.DataRegion,com.zhuozhengsoft.pageoffice.wordwriter.DataRegionInsertType,com.zhuozhengsoft.pageoffice.wordwriter.WordDocument,javax.servlet.*"
         pageEncoding="utf-8" %>
<%@ page import="javax.servlet.http.*" %>
<%
    if (request.getParameter("ids").equals(null)
            || request.getParameter("ids").equals("")) {
        return;
    }
    String idlist = request.getParameter("ids").trim();
    String[] ids = idlist.split(","); //将idlist按照","截取后存到ids数组中，然后遍历数组用js插入文件
    String temp = "PO_begin";//存储数据区域名称
    int num = 1;//试题编号
    WordDocument doc = new WordDocument();
    for (int i = 0; i < ids.length; i++) {

        DataRegion dataNum = doc.createDataRegion("PO_" + num,
                DataRegionInsertType.After, temp);
        dataNum.setValue(num + ".\t");
        DataRegion dataRegion = doc.createDataRegion("PO_begin"
                + (i + 1), DataRegionInsertType.After, "PO_" + num);
        dataRegion.setValue("[word]Openfile.jsp?id=" + ids[i]
                + "[/word]");
        temp = "PO_begin" + (i + 1);
        num++;
    }

    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    poCtrl1.setCustomToolbar(false);
    poCtrl1.setCaption("生成试卷");
    poCtrl1.setWriter(doc);
    //打开Word文件
    poCtrl1.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>

    <title>在Word文档中动态生成 试卷</title>

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
