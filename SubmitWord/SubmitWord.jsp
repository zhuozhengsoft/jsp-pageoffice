<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.DataRegion,com.zhuozhengsoft.pageoffice.wordwriter.WordDocument"
         pageEncoding="utf-8" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    WordDocument wordDoc = new WordDocument();
    //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
    DataRegion dataRegion1 = wordDoc.openDataRegion("PO_userName");
    //设置DataRegion的可编辑性
    dataRegion1.setEditing(true);
    //为DataRegion赋值,此处的值可在页面中打开Word文档后自己进行修改
    dataRegion1.setValue("");

    DataRegion dataRegion2 = wordDoc.openDataRegion("PO_deptName");
    dataRegion2.setEditing(true);
    dataRegion2.setValue("");

    poCtrl.setWriter(wordDoc);

    //添加自定义按钮
    poCtrl.addCustomToolButton("保存", "Save", 1);
    //设置保存页面
    poCtrl.setSaveDataPage("SaveData.jsp");
    //打开Word文档
    poCtrl.webOpen("doc/test.doc", OpenModeType.docSubmitForm, "张佚名");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>最简单的提交Word中的数据</title>
    <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
        }
    </script>
    <script type="text/javascript">
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
</head>
<body>
<form id="form1">
    <div>
        <span style="color: Red; font-size: 14px;">请输入公司名称、年龄、部门等信息后，单击工具栏上的保存按钮</span>
        <br/>
        <span style="color: Red; font-size: 14px;">请输入公司名称：</span>
        <input id="txtCompany" name="txtCompany" type="text"/>
        <br/>
    </div>
    <div style="width: auto; height: 700px;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>
