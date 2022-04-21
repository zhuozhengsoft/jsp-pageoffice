<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    //获取token值
    String mytoken=request.getHeader("Authorization");
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.addCustomToolButton("保存", "Save()", 1);
    poCtrl.addCustomToolButton("打印", "PrintFile()", 6);
    poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen()", 4);
    poCtrl.addCustomToolButton("关闭", "CloseFile()", 21);
    //设置保存页面
    poCtrl.setSaveFilePage("SaveFile.jsp");
    //打开Word文档
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张佚名");
%>
<!DOCTYPE html>
<html>
<head lang="en">
    <meta charset="UTF-8">
    <title>文件显示页</title>
    <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();

        }

        function PrintFile() {
            document.getElementById("PageOfficeCtrl1").ShowDialog(4);

        }

        function IsFullScreen() {
            document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;

        }

        function CloseFile() {
            window.external.close();
        }
    </script>
</head>
<body>
<h4>文件打开页面的token的值为：<%=mytoken %></h4>
    <div style="height: 900px;width: auto;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</body>
</html>

