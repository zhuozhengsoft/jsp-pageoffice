<%@ page language="java" import="com.zhuozhengsoft.pageoffice.wordreader.WordDocument" pageEncoding="utf-8" %>
<%@ page import="javax.servlet.*,javax.servlet.http.*" %>
<%
    String ErrorMsg = "";
    String BaseUrl = "";

    WordDocument doc = new WordDocument(request, response);
    String sName = doc.openDataRegion("PO_name").getValue();
    String sDept = doc.openDataRegion("PO_dept").getValue();
    String sCause = doc.openDataRegion("PO_cause").getValue();
    String sNum = doc.openDataRegion("PO_num").getValue();
    String sDate = doc.openDataRegion("PO_date").getValue();

    if (sName.equals("")) {
        ErrorMsg = ErrorMsg + "<li>申请人</li>";
    }
    if (sDept.equals("")) {
        ErrorMsg = ErrorMsg + "<li>部门名称</li>";
    }
    if (sCause.equals("")) {
        ErrorMsg = ErrorMsg + "<li>请假原因</li>";
    }
    if (sDate.equals("")) {
        ErrorMsg = ErrorMsg + "<li>日期</li>";
    }
    try {
        if (sNum != "") {
            if (Integer.parseInt(sNum) < 0) {
                ErrorMsg = ErrorMsg + "<li>请假天数不能是负数</li>";
            }
        } else {
            ErrorMsg = ErrorMsg + "<li>请假天数</li>";
        }
    } catch (Exception Ex) {
        ErrorMsg = ErrorMsg + "<li><font color=red>注意：</font>请假天数必须是数字</li>";
    }

    if (ErrorMsg == "") {
        // 您可以在此编程，保存这些数据到数据库中。
        out.println("提交的数据为：<br/>");
        out.println("姓名：" + sName + "<br/>");
        out.println("部门：" + sDept + "<br/>");
        out.println("原因：" + sCause + "<br/>");
        out.println("天数：" + sNum + "<br/>");
        out.println("日期：" + sDate + "<br/>");
        doc.showPage(578, 380);
    } else {
        ErrorMsg = "<div style='color:#FF0000;'>请修改以下信息：</div> "
                + ErrorMsg;
        doc.showPage(578, 380);
    }
    doc.close();
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<HTML>
<HEAD>
    <title>SaveData</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="C#" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
</HEAD>
<body>
<div class="errTopArea" style="TEXT-ALIGN:left">[提示标题：这是一个开发人员可自定义的对话框]</div>
<div class="errMainArea" id="error163">
    <div class="errTxtArea" style="HEIGHT:150px; TEXT-ALIGN:left">
        <b class="txt_title">
            <ul>
                <%=ErrorMsg%>
            </ul>
        </b>
    </div>

</div>
</body>
</HTML>

