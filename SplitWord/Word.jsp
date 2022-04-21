<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.DataRegion"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>

<%
    //******************************卓正PageOffice组件的使用*******************************
    PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
    poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    WordDocument wordDoc = new WordDocument();
    //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
    DataRegion dataRegion1 = wordDoc.openDataRegion("PO_test1");
    dataRegion1.setSubmitAsFile(true);
    DataRegion dataRegion2 = wordDoc.openDataRegion("PO_test2");
    dataRegion2.setSubmitAsFile(true);
    dataRegion2.setEditing(true);
    DataRegion dataRegion3 = wordDoc.openDataRegion("PO_test3");
    dataRegion3.setSubmitAsFile(true);

    poCtrl1.setWriter(wordDoc);
    poCtrl1.addCustomToolButton("保存", "Save()", 1);
    poCtrl1.setSaveDataPage("SaveData.jsp");
    //打开Word文件
    poCtrl1.webOpen("doc/test.doc", OpenModeType.docSubmitForm, "张三");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title></title>
</head>
<body>
<form id="form1">
    <div style="width: auto; height: 700px;">
        <!-- *********************PageOffice组件客户端JS代码*************************** -->
        <script type="text/javascript">
            function Save() {
                document.getElementById("PageOfficeCtrl1").WebSave();
            }
        </script>
        <div style=" font-size:14px; line-height:20px;">演示说明：<br/>点击“保存”按钮，PageOffice会把文档中三个数据区域（PO_test1，PO_test2，PO_test3）中的内容保存为三个独立的子文件（new1.doc，new2.doc，new3.doc）到“Samples4/SplitWord/doc”
            目录下。
        </div>
        <div style="color: red;font-size:14px; line-height:20px;">
            Word拆分功能只有企业版支持，并且文档的打开模式必须是OpenModeType.docSubmitForm，需要设置数据区域的属性dataRegion1.setSubmitAsFile(true)
            。<br/><br/></div>

        <!-- *********************PageOffice组件的引用*************************** -->
        <%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>
    </div>
    <img src="OpenStream.jsp"/>
</form>
</body>
</html>
