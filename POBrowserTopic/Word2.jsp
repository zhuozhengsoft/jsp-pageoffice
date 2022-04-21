<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>

<%
    //获取index.jsp页面传递过来参数的值
    String userName = (String) session.getAttribute("userName");
    //获取index.jsp用？传递过来的id的值
    String id = request.getParameter("id");
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    //添加自定义按钮
    poCtrl.addCustomToolButton("保存", "Save", 1);
    //设置保存页面
    poCtrl.setSaveFilePage("SaveFile.jsp");
    //打开Word文档
    //poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>获取index.jsp页面传递过来的参数</title>
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }
</script>
(1)获取index.jsp用session传递过来的userName的值：</br>
&nbsp;&nbsp;&nbsp;代码：String userName=(String)session.getAttribute("userName");</br>
&nbsp;&nbsp;&nbsp;输出：UserName=<%=userName %></br></br>

(2)获取index.jsp用？传递过来的id的值：</br>
&nbsp;&nbsp;&nbsp;代码：String id=request.getParameter("id");</br>
&nbsp;&nbsp;&nbsp;输出：id=<%=id %></br>
<form id="form1">
    <div style=" width:auto; height:700px;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>
