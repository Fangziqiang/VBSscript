<%@ page contentType="text/html; charset=utf-8"%>
<%@ page import="java.io.*"%>
<%@ page import="java.sql.*, javax.sql.*"%>
<%@ page import="java.util.*"%>
<%@ page import="java.math.*"%>
<HTML>
<HEAD>
<TITLE>图像测试</TITLE>

</HEAD>

<BODY>
	<%
		//String photo_no = request.getParameter("photo_no");

		//mysql连接
		Class.forName("com.mysql.jdbc.Driver").newInstance();
		String URL = "jdbc:mysql://localhost/ymc_touzijituan20190125?user=root&password=yascn&useUnicode=true&characterEncoding=utf8";
		Connection con = DriverManager.getConnection(URL);
	%>
	<TABLE align=center>
		<TR >
			<TD>plid</TD>
			<TD>姓名</TD>
		</TR>
		<%
			try {
				// 准备语句执行对象
				Statement stmt = con.createStatement();

				//String sql = " SELECT * FROM pltemplate WHERE id = "+ photo_no;
				String sql = "SELECT * from pltemplate where plid = '1561'";
				ResultSet rs = stmt.executeQuery(sql);

				while (rs.next()) {
		%>
		<TR border=1>
			<TD><%=rs.getInt("plid")%></TD>
			<TD><%=rs.getString("plname")%></TD>
			<TD><img src='show.jsp?id=<%=rs.getInt("plid")%>'></TD>
			 
			<TD><img src='show2.jsp?id=<%=rs.getInt("plid")%>'></TD>
			 
		</TR>
		<%
			}
				rs.close();
				con.close();
			} finally {

			}
		%>
	</TABLE>

</BODY>
</HTML>
