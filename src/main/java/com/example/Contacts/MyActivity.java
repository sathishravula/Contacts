package com.example.Contacts;

import android.app.Activity;
import android.content.ContentValues;
import android.database.Cursor;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.provider.Contacts;
import android.provider.ContactsContract;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

public class MyActivity extends Activity {
    HashMap<String,String> map;
    Button importContacts;
    Button exportContacts;
    /**
     * Called when the activity is first ctreated.
     */
    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.main);
       // Cursor cursor = getContentResolver().query(   ContactsContract.CommonDataKinds.Phone.CONTENT_URI, null, null,null, null);
        exportContacts= (Button) findViewById(R.id.button2);
        importContacts= (Button) findViewById(R.id.button);
        exportContacts.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                map=readContactsFromPhone();
                createSheet(map);
            }
        });
        importContacts.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                HashMap<String,String> data=null;
                try {
                    data=  readFromSheet();
                } catch (IOException e) {
                    e.printStackTrace();
                }
                Set entrySet=  data.entrySet();
                Iterator iterator=entrySet.iterator();
                while (iterator.hasNext()){
                    Map.Entry entry=(Map.Entry)iterator.next();
                    addContact((String)entry.getValue(),(String)entry.getKey());
                }
            }
        });
        //map=readContactsFromPhone();
        //createSheet(map);
/*
HashMap<String,String> data=null;
        try {
        data=  readFromSheet();
        } catch (IOException e) {
            e.printStackTrace();
        }
        Set entrySet=  data.entrySet();
        Iterator iterator=entrySet.iterator();
        while (iterator.hasNext()){
            Map.Entry entry=(Map.Entry)iterator.next();
            addContact((String)entry.getValue(),(String)entry.getKey());
        }
*/
    }



    private void addContact(String name, String phone) {
        ContentValues values = new ContentValues();
       // values.put(Contacts.People.NUMBER, phone);
        values.put(Contacts.People.STARRED,0);
        values.put(Contacts.People.NAME, name);
        Uri dataUri = getContentResolver().insert(Contacts.People.CONTENT_URI, values);
        Uri updateUri = Uri.withAppendedPath(dataUri, Contacts.People.Phones.CONTENT_DIRECTORY);
        values.clear();
        values.put(Contacts.People.Phones.TYPE, Contacts.People.TYPE_MOBILE);
        values.put(Contacts.People.NUMBER, phone);
        updateUri = getContentResolver().insert(updateUri, values);
    }

    public HashMap<String, String> readFromSheet() throws IOException {
        HashMap<String,String> resultSet=new HashMap<String, String>();

           Workbook w;
            try {
                //Workbook.getWorkbook(getAssets().open("Books1.xls"));
                w = Workbook.getWorkbook(getAssets().open("Book1.xls"));
                // Get the first sheet
                Sheet sheet = w.getSheet(0);
                // Loop over column and lines
                for (int j = 0; j < sheet.getRows(); j++) {
                    Cell nameCel = sheet.getCell(0, j);
                    Cell numberCel = sheet.getCell(1, j);
                    if(nameCel.getContents()!="" && numberCel.getContents()!="" ) {
                        Log.d("12345",nameCel.getContents()+"   "+numberCel.getContents());

                        resultSet.put(numberCel.getContents(), nameCel.getContents());
                    }
                        /*for (int i = 0; i < sheet.getColumns(); i++) {
                            Cell cel = sheet.getCell(i, j);
                            Log.d("12345", cel.getContents()+i+"");
                        }*/
                }
            } catch (BiffException e) {
                e.printStackTrace();
                Log.d("12345","exception");
            } catch (Exception e) {
                Log.d("12345","exception1");
                Log.d("12345",e.getMessage());
                Log.d("12345",e.getCause()+"");
                e.printStackTrace();
            }
        return resultSet;
    }
 HashMap<String, String> readContactsFromPhone() {
     HashMap<String,String> resultSet=new HashMap<String, String>();

     Cursor cursor = getContentResolver().query(ContactsContract.CommonDataKinds.Phone.CONTENT_URI, null, null, null, null);

       while (cursor.moveToNext()) {
        String name = cursor.getString(cursor.getColumnIndex(ContactsContract.CommonDataKinds.Phone.DISPLAY_NAME));
        String phoneNumber = cursor.getString(cursor.getColumnIndex(ContactsContract.CommonDataKinds.Phone.NUMBER));

           Log.d("data","Name: "+name+"   Number: "+phoneNumber);
           resultSet.put(phoneNumber,name);

    }
     return resultSet;
 }
    void createSheet(HashMap<String, String> map){
        try {
            File newFolder = new File(Environment.getExternalStorageDirectory(), "TestFolder");
            Log.d("12345",Environment.getExternalStorageDirectory().getAbsolutePath());
            if (!newFolder.exists()) {
                newFolder.mkdir();
                Log.d("12345", newFolder.getAbsolutePath());
                Log.d("12345","folder created");
            }
            try {
               // File file = new File(newFolder, "MyTest" + ".txt");
                //file.createNewFile();

                String fileName = "file.xls";
                WritableWorkbook workbook = Workbook.createWorkbook(new File(newFolder, "file.xls"));
                WritableSheet sheet = workbook.createSheet("Sheet1", 0);
                Set entrySet=  map.entrySet();
                Iterator iterator=entrySet.iterator();
                int i=0;
                while (iterator.hasNext()) {
                    Map.Entry entry = (Map.Entry) iterator.next();
                    //    addContact((String)entry.getValue(),(String)entry.getKey());

                    Label label = new Label(0, i, (String) entry.getValue());
                    Label label1 = new Label(1, i, (String) entry.getKey());
                    sheet.addCell(label);
                    sheet.addCell(label1);
                    i++;
                }

                workbook.write();
                workbook.close();

                Log.d("12345","file created");
            } catch (Exception ex) {
                System.out.println("ex: " + ex);
                Log.d("12345","file created1");
            }
        } catch (Exception e) {
            System.out.println("e: " + e);
            Log.d("12345","file created2");
        }

    }

}
