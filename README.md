# jsp-pageoffice
### 一、简介

​       jsp-pageoffice项目演示了在JSP下如何使用PageOffice产品，此项目演示了PageOffice产品近90%的功能，是一个demo项目。

### 二、项目环境要求

​     jdk1.8及以上版本，tomcat8.0及以上版本

### 三、项目运行步骤

1. 使用git clone或者直接下载项目压缩包到本地并解压缩。
2. 拷 贝 jsp-pageoffice文 件 夹 到Tomcat的Webapps目录下,启动tomcat。
3. 启动浏览器访问：http://localhost:8080/jsp-pageoffice/index.html ，即可在线打开、编辑、保存office文件 。

### 四、PageOffice序列号

​     PageOfficeV5.0标准版试用序列号：I2BFU-MQ89-M4ZZ-ZWY7K           
​     PageOfficeV5.0专业版试用序列号：DJMTF-HYK4-BDQ3-2MBUC

### 五、集成PageOffice到您的项目中的关键步骤

1. 拷贝“jsp-pageoffice/lib”文件夹中的pagoffice5.x.x.x.jar拷贝到您自己Web项目中WEB-INF目录下；
2. 参考“jsp-pageoffice/WEB-INF”文件夹中的web.xml文件，配置一下自己Web项目中的web.xml

文件，也可以直接把下面的这段配置添加到自己的web.xml 中；

```xml
<!-- PageOffice Begin -->
<servlet>
	<servlet-name>poserver</servlet-name>
	<servlet-class>com.zhuozhengsoft.pageoffice.poserver.Server</servlet-class>
</servlet>
<servlet-mapping>
	<servlet-name>poserver</servlet-name>
	<url-pattern>/poserver.zz</url-pattern>
</servlet-mapping>
<servlet-mapping>
	<servlet-name>poserver</servlet-name>
	<url-pattern>/sealsetup.exe</url-pattern>
</servlet-mapping>
<servlet-mapping>
	<servlet-name>poserver</servlet-name>
	<url-pattern>/posetup.exe</url-pattern>
</servlet-mapping>
<servlet-mapping>
	<servlet-name>poserver</servlet-name>
	<url-pattern>/pageoffice.js</url-pattern>
</servlet-mapping>
<servlet-mapping>
	<servlet-name>poserver</servlet-name>
	<url-pattern>/jquery.min.js</url-pattern>
</servlet-mapping>
<servlet-mapping>
	<servlet-name>poserver</servlet-name>
	<url-pattern>/pobstyle.css</url-pattern>
</servlet-mapping>
<servlet>
	<servlet-name>adminseal</servlet-name>
	<servlet-class>com.zhuozhengsoft.pageoffice.poserver.AdminSeal</servlet-class>
</servlet>
<servlet-mapping>
	<servlet-name>adminseal</servlet-name>
	<url-pattern>/adminseal.zz</url-pattern>
</servlet-mapping>
<servlet-mapping>
	<servlet-name>adminseal</servlet-name>
	<url-pattern>/loginseal.zz</url-pattern>
</servlet-mapping>
<servlet-mapping>
	<servlet-name>adminseal</servlet-name>
	<url-pattern>/sealimage.zz</url-pattern>
</servlet-mapping>
<mime-mapping>
	<extension>mht</extension>
	<mime-type>message/rfc822</mime-type>
</mime-mapping>
<!--为了安全，强烈建议修改此密码-->
<context-param>
    <param-name>adminseal-password</param-name>
    <param-value>111111</param-value>
</context-param>
<!-- PageOffice End -->
```
3. 在需要在线编辑的 JSP 页面头部添加：

  `<%@ page language="java" import="com.zhuozhengsoft.pageoffice.*" %>`

4. 对 PageOffice 编程控制：

```java
<%
  PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request); 
  poCtrl1.setServerPage("poserver.zz"); //此行必须
  poCtrl1.setSaveFilePage("savefile.jsp");//如要保存文件，此行必须
  //打开文件
  poCtrl1.webOpen("doc/test.doc", OpenModeType.docNormalEdit, "张三");
%>
```

5. 在 html 代码的 body 中需要出现 office 界面的位置插入下面的代码：

`<%=poCtrl1.getHtmlCode("PageOfficeCtrl1")%>`

6. 如果使用 PageOffice 的 POBrowser 方式打开文件， 那么调用 javascript 方法

“POBrowser.openWindowModeless”的页面一定要引用下面的 js 文件：

`<script type="text/javascript" src="/pageoffice.js"></script>`

> 【注意】：pageoffice.js 文件的位置在第 2 步配置web.xml 的时候已经设置好了，
>
> 所以直接引用这个 js 即可，无需拷贝 pageoffice.js 文件到自己的Web项目目录下。


### 六、电子印章功能说明

如果您用到电子印章功能，请注意以下事项：

1. 您部署网站时，只需修改 web.xml 中的以下内容，其他PageOffice节点内容都无需更改。

动。

```xml
<context-param>
  <param-name>adminseal-password</param-name>
  <!--这里修改成您的印章简易管理页的管理员登录密码，为了安全，强烈建议修改!!!-->
  <param-value>111111</param-value>
</context-param>
```

2. 印章数据库默认的SQLite数据库，如果您需要改用Oracle、MySQL等其他数据库,请将

下面的配置添加到web.xml 中,并且修改成您的印章数据库连接参数。最好把 WEB-INF\lib 下 

的 poseal.db 也删除掉。例如连接oracle数据库：

```xml
<context-param>
  <param-name>posealdb-driver</param-name>
  <param-value>oracle.jdbc.driver.OracleDriver</param-value>
</context-param>
<context-param>
  <param-name>posealdb-url</param-name>
  <param-value>jdbc:oracle:thin:@192.168.0.1:1521:orcl</param-value>
</context-param>
<context-param>
  <param-name>posealdb-username</param-name>
  <param-value>system</param-value>
</context-param>
<!-- 为了安全，强烈建议修改此密码--> 
<context-param>
  <param-name>posealdb-password</param-name>
  <param-value>111111</param-value>
</context-param>
```

3. 请参考当前项目的一、15、演示加盖印章和签字功能（以Word为例）http://localhost:8080/jsp-pageoffice/InsertSeal/index.jsp  示例代码进行盖章功能调用。

### 七、 PageOffice开发帮助

​     1.[Java API文档](https://www.zhuozhengsoft.com/help/java3/index.html) 

​     2.[JS API文档](https://www.zhuozhengsoft.com/help/js3/index.html)  

​     3.[PageOffice从入门到精通](https://www.kancloud.cn/pageoffice_course_group/pageoffice_course/646953)

​     技术支持：https://www.zhuozhengsoft.com/Technical/

### 八、联系我们

​   卓正官网：[https://www.zhuozhengsoft.com](http://www.zhuozhengsoft.com)

​   联系电话：400-6600-770  

   QQ: 800038353
