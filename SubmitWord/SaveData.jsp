<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.wordreader.DataRegion,com.zhuozhengsoft.pageoffice.wordreader.WordDocument"
         pageEncoding="utf-8" %>
<%
    WordDocument doc = new WordDocument(request, response);
    //获取提交的数值
    DataRegion dataUserName = doc.openDataRegion("PO_userName");
    DataRegion dataDeptName = doc.openDataRegion("PO_deptName");
    String content = "";
    content += "公司名称：" + doc.getFormField("txtCompany");
    content += "<br/>员工姓名：" + dataUserName.getValue();
    content += "<br/>部门名称：" + dataDeptName.getValue();
    doc.showPage(500, 400);
    doc.close();
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title></title>
</head>
<body>
<form id="form1">
    <div style="border: solid 1px gray;">
        <div class="errTopArea"
             style="text-align: left; border-bottom: solid 1px gray;">
            [提示标题：这是一个开发人员可自定义的对话框]
        </div>
        <div class="errTxtArea" style="height: 150px; text-align: left">
            <b class="txt_title">
                <div style="color: #FF0000;">
                    提交的信息如下：
                </div>
                <%=content%>
            </b>
        </div>
        <div class="errBtmArea" style="text-align: center;">
            <input type="button" class="btnFn" value=" 关闭 "
                   onclick="window.opener=null;window.open('','_self');window.close();"/>
        </div>
    </div>
</form>
</body>
</html>