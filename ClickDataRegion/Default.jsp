<%@ page language="java" import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<%@ page import="com.zhuozhengsoft.pageoffice.wordwriter.DataRegion,com.zhuozhengsoft.pageoffice.wordwriter.WordDocument" %>
<%@ page import="java.awt.*" %>
<%
    PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);

    WordDocument doc = new WordDocument();
    DataRegion dataReg = doc.openDataRegion("PO_deptName");
    dataReg.getShading().setBackgroundPatternColor(Color.pink);
    //dataReg.setEditing(true);
    poCtrl.setWriter(doc);

    //设置服务器页面
    poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
    poCtrl.addCustomToolButton("保存", "Save", 1);
    poCtrl.setJsFunction_OnWordDataRegionClick("OnWordDataRegionClick()");
    poCtrl.setOfficeToolbars(false);
    poCtrl.setCaption("为方便用户知道哪些地方可以编辑，所以设置了数据区域的背景色");
    poCtrl.setSaveFilePage("SaveFile.jsp");
    poCtrl.webOpen("doc/test.doc", OpenModeType.docSubmitForm, "张三");
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
    <div id="textcontent" style="width: 1000px; height: 800px;">
        <script type="text/javascript">
            //***********************************PageOffice组件调用的js函数**************************************
            //保存页面
            function Save() {
                document.getElementById("PageOfficeCtrl1").WebSave();
            }

            //全屏/还原
            function IsFullScreen() {
                document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
            }

            function OnWordDataRegionClick(Name, Value, Left, Bottom) {

                if (Name == "PO_deptName") {
                    var strRet = document.getElementById("PageOfficeCtrl1").ShowHtmlModalDialog("HTMLPage.htm", Value, "left=" + Left + "px;top=" + Bottom + "px;width=400px;height=300px;frame=no;");
                    if (strRet != "") {
                        return (strRet);
                    }
                    else {
                        if ((Value == undefined) || (Value == ""))
                            return " ";
                        else
                            return Value;
                    }
                }
            }
        </script>
        <!--**************   卓正 PageOffice组件 ************************-->
        <%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
</div>
</body>
</html>
