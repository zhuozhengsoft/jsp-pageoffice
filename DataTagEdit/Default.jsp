<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.BorderStyleType,com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.ThemeType" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>
<%
    WordDocument doc = new WordDocument();
    doc.getTemplate().defineDataTag("{ 甲方 }");
    doc.getTemplate().defineDataTag("{ 乙方 }");
    doc.getTemplate().defineDataTag("{ 担保人 }");
    doc.getTemplate().defineDataTag("【 合同日期 】");
    doc.getTemplate().defineDataTag("【 合同编号 】");

    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
    poCtrl.addCustomToolButton("保存", "Save()", 1);
    poCtrl.addCustomToolButton("定义数据标签", "ShowDefineDataTags()", 20);
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
    poCtrl.setSaveFilePage("SaveFile.jsp");
    poCtrl.setTheme(ThemeType.Office2007);
    poCtrl.setBorderStyle(BorderStyleType.BorderThin);
    poCtrl.setWriter(doc);
    poCtrl.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "zhangsan");
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
    <title></title>
    <script type="text/javascript">
        //获取后台定义的Tag 字符串
        function getTagNames() {
            var tagNames = document.getElementById("PageOfficeCtrl1").DefineTagNames;
            return tagNames;
        }

        var m_strIsChange = "no";//标识是否更换了要定位的标签
        //更换了要定位的数据标签
        function changeTag(strIsChange) {
            m_strIsChange = strIsChange;
        }

        //定位Tag
        function locateTag(tagName) {
            var appSlt = document.getElementById("PageOfficeCtrl1").Document.Application.Selection;
            var bFind = false;
            //更换了要定位的标签时，先回到文档开始位置再进行定位
            if (m_strIsChange == "yes"){
                appSlt.HomeKey(6);
            }
            appSlt.Find.ClearFormatting();
            appSlt.Find.Replacement.ClearFormatting();
            bFind = appSlt.Find.Execute(tagName);
            if (!bFind) {
                document.getElementById("PageOfficeCtrl1").Alert("已搜索到文档末尾。");
                //appSlt.HomeKey(6);
            }
            window.focus();
        }

        //添加Tag
        function addTag(tagName) {
            try {
                var tmpRange = document.getElementById("PageOfficeCtrl1").Document.Application.Selection.Range;
                tmpRange.Text = tagName;
                tmpRange.Select();
                return "true";
            } catch (e) {
                return "false";
            }
        }

        //删除Tag
        function delTag(tagName) {
            var tmpRange = document.getElementById("PageOfficeCtrl1").Document.Application.Selection.Range;
            if (tagName == tmpRange.Text) {
                tmpRange.Text = "";
                return "true";
            }
            else
                return "false";
        }
    </script>
    <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
        }

        function ShowDefineDataTags() {
            document.getElementById("PageOfficeCtrl1").ShowHtmlModelessDialog("DataTagDlg.htm", "parameter=xx", "left=300px;top=390px;width=430px;height=410px;frame:no;");
        }
    </script>
</head>
<body>
<form action="">
    <div style="width: auto; height: 600px;">
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</form>
</body>
</html>