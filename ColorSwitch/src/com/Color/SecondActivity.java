package com.Color;
import android.app.Activity;
import android.os.Bundle;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.widget.EditText;
import android.widget.TextView;

public class SecondActivity extends Activity {

	EditText mEdit;
	String data=null;
	TextView textData;
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_second);
	}
	
	

	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		// Inflate the menu; this adds items to the action bar if it is present.
		getMenuInflater().inflate(R.menu.second, menu);
		return true;
	}

	@Override
	public boolean onOptionsItemSelected(MenuItem item) {
		// Handle action bar item clicks here. The action bar will
		// automatically handle clicks on the Home/Up button, so long
		// as you specify a parent activity in AndroidManifest.xml.
		int id = item.getItemId();
		if (id == R.id.action_settings) {
			return true;
		}
		return super.onOptionsItemSelected(item);
	}
	
	public void displayData(View view){
		mEdit   = (EditText)findViewById(R.id.name);
		data=mEdit.getText().toString();
		
		mEdit = (EditText)findViewById(R.id.email);
		data=data.concat("\n".concat(mEdit.getText().toString()));
		
		mEdit = (EditText)findViewById(R.id.phone);
		data=data.concat("\n".concat(mEdit.getText().toString()));
		
		textData=(TextView) findViewById(R.id.TextView);
		textData.setText(data);
		
	}
}
