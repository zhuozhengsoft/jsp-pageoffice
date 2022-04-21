<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    String userName = "somebody";
    String userId = request.getParameter("userid").toString();
    if (userId.equals("1")) {
        userName = "张三";
    } else {
        userName = "李四";
    }

    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.addCustomToolButton("保存", "Save", 1);
    poCtrl.setSaveFilePage("SaveFile.jsp");
    //设置并发控制时间
    poCtrl.setTimeSlice(20);
    poCtrl.webOpen("doc/test.doc", OpenModeType.docRevisionOnly, userName);
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>最简单的打开保存Word文件</title>
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }

    //文档关闭前先提示用户是否保存
    function BeforeBrowserClosed() {
        if (document.getElementById("PageOfficeCtrl1").IsDirty) {
            if (confirm("提示：文档已被修改，是否继续关闭放弃保存 ？")) {
                return true;

            } else {

                return false;
            }

        }

    }
</script>
<form id="form1">
    当前用户： <%=userName %>。
    <div style=" width:auto; height:700px;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>
