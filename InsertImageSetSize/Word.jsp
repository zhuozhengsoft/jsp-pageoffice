<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.DataRegion"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>
<%
    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    WordDocument worddoc = new WordDocument();
    //先在要插入word文件的位置手动插入书签,书签必须以“PO_”为前缀
    //给DataRegion赋值,值的形式为："[word]word文件路径[/word]、[excel]excel文件路径[/excel]、[image]图片路径[/image]"
    DataRegion data1 = worddoc.openDataRegion("PO_p1");
    // [image]标签中的width和height属性的单位是：英磅
    data1.setValue("[image width=200.2 height=200]doc/1.jpg[/image]");
    poCtrl1.setWriter(worddoc);
    poCtrl1.setCaption("演示：后台编程插入图片到数据区域并设置图片大小(企业版)");
    //隐藏菜单栏
    poCtrl1.setMenubar(false);
    //隐藏自定义工具栏
    poCtrl1.setCustomToolbar(false);
    poCtrl1.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>演示：后台编程插入图片到数据区域(专业版、企业版)</title>
</head>
<body>
<div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc;
        padding: 5px;">
    关键代码：<span style="background-color: Yellow;"> <br/>DataRegion dataRegion
            = worddoc.openDataRegion("PO_开头的书签名称");
            <br/>
		dataRegion.setValue("[image width=200.2 height=200]doc/1.jpg[/image]");//[image]标签中的width和height属性的单位是：英磅</span><br/>
</div>
<br/>
<form id="form1" style="height: 100%;">
    <div style="height: 700px; width: auto;">
        <!--**************   PageOffice 客户端代码开始    ************************-->
        <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
        <!--**************   PageOffice 客户端代码结束    ************************-->
    </div>
</form>
</body>
</html>
