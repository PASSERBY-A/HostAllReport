package sdhostall;

import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Properties;

public class DBManager {

	private static DBManager instance;
	
	
	private DBManager()
	{
	}
	
	private static Connection connection = null;
	
	private static Properties prop = null;
	static{
		try {
			
			 prop = new Properties();
			
			prop.load(DBManager.class.getResourceAsStream("/config.properties"));
			
			Class.forName(prop.getProperty("jdbc.connection.driverclass"));
			
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}   catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
	
	
	public  Connection getConnection()
	{
		
		try {
			if(connection==null)
			{
				connection = DriverManager.getConnection(prop.getProperty("jdbc.connection.url"),prop.getProperty("jdbc.connection.username"),prop.getProperty("jdbc.connection.password"));
				
				return connection;
			}
			
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return connection;
		
	}
	
	public void release(Connection connection,Statement smt,ResultSet rs) throws SQLException
	{
		if(connection!=null)
		{
				connection.close();
				connection =null;
				
		}
		
		if(smt!=null)
		{
			
			smt.close();
			smt=null;
		}
		
		if(rs!=null)
		{
			rs.close();
			rs =null;
		}
		
		
	}
	
	
	
	
	
	public static DBManager getInstance()
	{
		if(instance==null)
		{
			
			instance = new DBManager();
			
		}
		
		return instance;
	}
	
}
