<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.DocumentOpenType,com.zhuozhengsoft.pageoffice.FileMakerCtrl"
         pageEncoding="utf-8" %>
<%
    String filePath = request.getSession().getServletContext()
            .getRealPath("InsertSeal/doc/");
    String id = request.getParameter("id").trim();
    if ("1".equals(id)) {
        filePath = "doc/test1.doc";
    }
    if ("2".equals(id)) {
        filePath = "doc/test2.doc";
    }
    if ("3".equals(id)) {
        filePath = "doc/test3.doc";
    }
    if ("4".equals(id)) {
        filePath = "doc/test4.doc";
    }

    FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
    fmCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    fmCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
    fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
    fmCtrl.setSaveFilePage("SaveFile.jsp");
    fmCtrl.fillDocument(filePath, DocumentOpenType.Word);
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>My JSP 'FileMaker.jsp' starting page</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
    <script type="text/javascript">
        //标识盖章是否成功
        var isAddSealSuc = false;

        function AfterDocumentOpened() {
            try {
                //先定位到印章位置,再在印章位置上盖章
                document.getElementById("FileMakerCtrl1").ZoomSeal.LocateSealPosition("Seal1");
                isAddSealSuc = document.getElementById("FileMakerCtrl1").ZoomSeal.AddSealByName("测试公章", null, "Seal1"); //位置名称不区分大小写
            } catch (e) {
            }
        }

        function OnProgressComplete() {
            window.parent.convert(isAddSealSuc); //调用父页面的js函数
        }
    </script>
</head>
<body>
<div>
    <%=fmCtrl.getHtmlCode("FileMakerCtrl1")%>
</div>
</body>
</html>