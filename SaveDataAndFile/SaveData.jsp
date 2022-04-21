<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.wordreader.WordDocument"
         pageEncoding="utf-8" %>
<%
    WordDocument doc = new WordDocument(request, response);
    //获取提交的数值
    String dataUserName = doc.openDataRegion("PO_userName").getValue();
    String dataDeptName = doc.openDataRegion("PO_deptName").getValue();
    String companyName = doc.getFormField("txtCompany");

    /**获取到的公司名称,员工姓名,部门名称等内容可以保存到数据库,这里可以连接开发人员自己的数据库,例如连接mysql数据库。
     *Class.forName("com.mysql.jdbc.Driver");
     *Connection conn = (Connection) DriverManager.getConnection("jdbc:mysql://localhost/myDataBase", "root", "111111");
     *String sql="update user set userName='"+dataUserName+"',deptName='"+dataDeptName+"',companyName='"+companyName+"' where userId=1";
     *PreparedStatement ps = conn.prepareStatement(sql);
     *int rs = ps.executeUpdate(upsql);
     * if (rs>0) {
     *    out.println("更新成功");
     *     }
     *     else{
     *   out.println("更新失败");
     *    }
     *
     *rs.close();
     *conn.close();
     */
    doc.close();
%>

