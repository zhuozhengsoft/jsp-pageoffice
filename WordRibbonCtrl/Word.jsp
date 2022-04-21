<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    //添加自定义按钮
    poCtrl.addCustomToolButton("保存", "Save", 1);
    poCtrl.getRibbonBar().setTabVisible("TabHome", true);//开始
    poCtrl.getRibbonBar().setTabVisible("TabPageLayoutWord", false);//页面布局
    poCtrl.getRibbonBar().setTabVisible("TabReferences", false);//引用
    poCtrl.getRibbonBar().setTabVisible("TabMailings", false);//邮件
    poCtrl.getRibbonBar().setTabVisible("TabReviewWord", false);//审阅
    poCtrl.getRibbonBar().setTabVisible("TabInsert", false);//插入
    poCtrl.getRibbonBar().setTabVisible("TabView", false);//视图
    poCtrl.getRibbonBar().setSharedVisible("FileSave", false);//office自带的保存按钮
    poCtrl.getRibbonBar().setGroupVisible("GroupClipboard", false);//分组剪贴板
    //设置保存页面
    //poCtrl.setSaveFilePage("SaveFile.jsp");
    //打开Word文档
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张佚名");
    //poCtrl.webOpen("doc/test.docx",OpenModeType.docNormalEdit,"张佚名");
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
</script>
<form id="form1">
    <div style=" width:auto; height:700px;">
        <div style="font-size:14px; line-height:20px; ">
            poCtrl.getRibbonBar().setTabVisible("TabHome",false);实现隐藏Ribbon栏中的“开始”命令分组.（"TabHome"为该命令分组的名称,false为隐藏，true为显示）
            <br/>
            poCtrl.getRibbonBar().setSharedVisible("FileSave",
            false);实现隐藏Ribbon快速工具栏中的“保存”按钮.（"FileSave"为按钮的名称,false为隐藏，true为显示）
            <br/>
            poCtrl.getRibbonBar().setGroupVisible("GroupClipboard",
            false);实现隐藏“开始”命令分组中的“剪切板”面板.（"GroupClipboard"为该面板的名称,false为隐藏，true为显示）
            <br/>
            <span style="color:red">注意：此控制是会同步影响到本地文件打开时Ribbon栏中的各个工具按钮的显示状态，当关闭在线编辑窗口时，本地Ribbon栏状态恢复正常。</span>
            <br/><br/>
        </div>
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>
