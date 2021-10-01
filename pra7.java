
import java.sql.*;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import com.agile.api.AgileSessionFactory;
import com.agile.api.IAgileSession;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;


public class session {
	// JDBC driver name and database URL
	static final String JDBC_DRIVER = "oracle.jdbc.driver.OracleDriver";
    static final String DB_URL = "jdbc:oracle:thin:@192.168.222.132:1521:agile9";

    //  Database credentials
    static final String USER = "agile";
    static final String PASS = "tartan";

	public static void main(String[] args) {
		Connection conn = null;
	    Statement stmt = null;

		try {
			AgileSessionFactory instance = AgileSessionFactory.getInstance("http://janice-anselm:7001/Agile/");
			HashMap params = new HashMap();
			params.put(AgileSessionFactory.USERNAME, "admin");
			params.put(AgileSessionFactory.PASSWORD, "agile936");;
			IAgileSession session = instance.createSession(params);
			System.out.println("連結成功");
			System.out.println("-----------");

			//STEP 2: Register JDBC driver
			Class.forName("oracle.jdbc.driver.OracleDriver");

			//STEP 3: Open a connection
			System.out.println("Connecting to database...");
			conn = DriverManager.getConnection(DB_URL, USER, PASS);

			//STEP 4: Execute a query
			System.out.println("Creating statement...");
			stmt = conn.createStatement();
			System.out.println(stmt);
			String sql;
			sql = "SELECT * FROM ITEM";
			ResultSet rs = stmt.executeQuery(sql);

			while(rs.next()){
				String num  = rs.getString("ITEM_NUMBER");
				System.out.println(num);
		    }


		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
