package com.Color;

import com.Color.R;
import android.os.Bundle;
import android.app.Activity;
import android.content.Intent;
import android.graphics.Color;
import android.view.Menu;
import android.view.View;

public class ColorActivity extends Activity {

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_color);
	}

	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		// Inflate the menu; this adds items to the action bar if it is present.
		getMenuInflater().inflate(R.menu.color, menu);
		return true;
	}
	 public void red(View v){
			View view = this.getWindow().getDecorView();
		    view.setBackgroundColor(Color.RED); 
		}
		
		public void blue(View v){
			View view = this.getWindow().getDecorView();
		    view.setBackgroundColor(Color.BLUE); 
		}
		
		public void green(View v){
			View view = this.getWindow().getDecorView();
		    view.setBackgroundColor(Color.GREEN); 
		}
		
		public void black(View v){
			View view = this.getWindow().getDecorView();
		    view.setBackgroundColor(Color.BLACK); 
		}
		
		public void formActivity(View V){
			Intent i=new Intent(this, SecondActivity.class);
			startActivity(i);
		}
}
