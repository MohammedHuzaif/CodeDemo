package com.Color.test;

import java.io.IOException;

import android.os.Environment;
import android.test.ActivityInstrumentationTestCase2;
import android.util.Log;
import android.widget.TextView;

import com.jayway.android.robotium.solo.Solo;

@SuppressWarnings("unchecked")

public class TestCaseOne extends ActivityInstrumentationTestCase2 {
	
	private static final String TARGET_PACKAGE_ID = "com.Color";
	private static final String LAUNCHER_ACTIVITY_FULL_CLASSNAME = "com.Color.SplashActivity";
	private static Class<?> launcherActivityClass;
	private static final String TAG = "MyApp";
	
	static{
		try	{	
			launcherActivityClass=Class.forName(LAUNCHER_ACTIVITY_FULL_CLASSNAME);
		}catch(ClassNotFoundException e){ 
			throw new RuntimeException(e);
			}
	}
	
	public TestCaseOne() throws ClassNotFoundException
	{
	super(TARGET_PACKAGE_ID,launcherActivityClass);	
	}

	private Solo solo;

	@Override protected void setUp()throws Exception
		{ 
		super.setUp();
		solo = new Solo(getInstrumentation(),getActivity());
		}

	public void testStart()throws Exception{

		Log.v(TAG,"*******************************");
		Log.v(TAG,"*******************************");
		Log.v(TAG, "start testStart()");
		
		Log.v(TAG, "Click on string 'Touch to Continue'");
		solo.clickOnText("Touch to Continue");
		
		Log.v(TAG, "Click button RED");
		solo.clickOnButton("RED");
		
		Log.v(TAG, "Click button BLACK");
		solo.clickOnButton("BLACK");
		
		Log.v(TAG, "Click button BLUE");
		solo.clickOnButton("BLUE");
		
		Log.v(TAG, "Click button GREEN");
		solo.clickOnButton("GREEN");
		
		Log.v(TAG, "Click button Form");
		solo.clickOnButton("Form");

		Log.v(TAG, "Get header text");
		
		
		//String header=solo.getText(0).toString();
		  TextView text = (TextView) solo.getView("textView1");

		  String header = text.getText().toString();
		
		if(header.equals("Form")){
			Log.v(TAG,"Header Text matches the requirement 'Form'");
		}else
			Log.v(TAG,"Header Text does not match with the requirement 'Form'");
		Log.v(TAG, "Header text on the screen is '"+header+"'");
		
		String name="This is Name";
		Log.v(TAG, "Enter text in name edit text '"+name+"'");
		solo.enterText(0, name);
		
		String email="email@company.com";
		Log.v(TAG, "Enter text in email edit text '"+email+"'");
		solo.enterText(1, email);
		
		String phone="1234567890";
		Log.v(TAG, "Enter text in Phone edit text '"+phone+"'");
		solo.enterText(2, phone);
		
		Log.v(TAG, "Click on Submit button");
		solo.clickOnButton("Submit");
		String submitString=name.concat("\n".concat(email.concat("\n".concat(phone))));
		
		Thread.sleep(2000);

		Log.v(TAG, "Get text which is displayed after submit");
		
		TextView text1 = (TextView) solo.getView("TextView");

		String data = text1.getText().toString();
		  
		if(submitString.equals(data)){
			Log.v(TAG,"Text matches the data given in the form");
		}else
			Log.v(TAG,"Text does not match the data given in the form");
		Log.v(TAG, "data displayed on the screen is '"+data+"'");
		
		//solo.pressMenuItem(0);
		
		Log.v(TAG,"End of Test Case");
		Log.v(TAG,"*******************************");
		Log.v(TAG,"*******************************");
		getLog();
		}
	
	public void getLog(){
		try {			
			String filePath1 = Environment.getExternalStorageDirectory() + "/logcat1.txt";
			Runtime.getRuntime().exec(new String[]{"logcat", "-v", "time", "-f", filePath1, "MyApp:V","*:S"});
		    } catch (IOException e) {
		  }
	}

	@Override
	public void tearDown() throws Exception
		{
		try
			{
			solo.finalize();
			}
		catch(Throwable e)
		{
			e.printStackTrace();
		}
		getActivity().finish();
		super.tearDown();
		}

}
