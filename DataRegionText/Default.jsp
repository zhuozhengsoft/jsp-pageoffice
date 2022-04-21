<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.DataRegion,com.zhuozhengsoft.pageoffice.wordwriter.WdParagraphAlignment" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>
<%@ page import="java.awt.*" %>
<%
    WordDocument doc = new WordDocument();
    DataRegion d1 = doc.openDataRegion("d1");
    d1.getFont().setColor(Color.BLUE);//设置数据区域文本字体颜色
    d1.getFont().setName("华文彩云");//设置数据区域文本字体样式
    d1.getFont().setSize(16);//设置数据区域文本字体大小
    d1.getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);//设置数据区域文本对齐方式
    DataRegion d2 = doc.openDataRegion("d2");
    d2.getFont().setColor(Color.orange);//设置数据区域文本字体颜色
    d2.getFont().setName("黑体");//设置数据区域文本字体样式
    d2.getFont().setSize(14);//设置数据区域文本字体大小
    d2.getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphLeft);//设置数据区域文本对齐方式
    DataRegion d3 = doc.openDataRegion("d3");
    d3.getFont().setColor(Color.magenta);//设置数据区域文本字体颜色
    d3.getFont().setName("华文行楷");//设置数据区域文本字体样式
    d3.getFont().setSize(12);//设置数据区域文本字体大小
    d3.getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphRight);//设置数据区域文本对齐方式

    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    poCtrl.setWriter(doc);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    //设置文档打开方式
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
    <title></title>
    <link href="images/csstg.css" rel="stylesheet" type="text/css"/>
</head>
<body>
<div id="content">
    演示了如果使用程序控制数据区域文本的样式。<a href="Default2.jsp">点此链接查看原文件内容样式</a><br/><br/>
    <div id="textcontent" style="width: 1000px; height: 800px;">

        <!--**************   卓正 PageOffice组件 ************************-->
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</div>
</body>
</html>
