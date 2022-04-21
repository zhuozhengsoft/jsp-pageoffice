<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType, com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl pocCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    pocCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    //添加自定义按钮
    pocCtrl.addCustomToolButton("保存", "Save()", 1);
    pocCtrl.addCustomToolButton("另存为PDF文件", "SaveAsPDF()", 1);
    //设置保存页面
    pocCtrl.setSaveFilePage("SaveFile.jsp");
    String fileName = "template.doc";
    String pdfName = fileName.substring(0, fileName.length() - 4) + ".pdf";
    //打开文件
    pocCtrl.webOpen("doc/" + fileName, OpenModeType.docNormalEdit, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>Word文件转换成PDF格式</title>
    <script type="text/javascript">
        //保存
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
        }

        //另存为PDF文件
        function SaveAsPDF() {
            document.getElementById("PageOfficeCtrl1").WebSaveAsPDF();
            document.getElementById("PageOfficeCtrl1").Alert("PDF文件已经保存到 SaveAsPDF\\doc目录下。");
            document.getElementById("div1").innerHTML = "<a href='OpenPDF.jsp?fileName=<%=pdfName %>'> 查看另存的 pdf 文件<a><br><br>";
        }
    </script>
</head>
<body>
<form id="form1">
    <div id="div1"></div>
    <div style="width: auto; height: 700px;">
        <%=pocCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>

