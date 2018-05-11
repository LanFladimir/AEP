package com.fladimir.aep;

import android.os.Bundle;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class MainActivity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        //declare 声明

        String cells[] = {"TitleA", "TitleB", "TitleC"};

        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        HSSFSheet sheetA = hssfWorkbook.createSheet("SheetA");
        {
            HSSFCell cell;
            HSSFRow row = sheetA.createRow(0);
            for (int i = 0; i < cells.length; i++) {
                cell = row.createCell(i);
                cell.setCellValue(cells[i]);
            }

            for (int i = 1; i < 10; i++) {
                row = sheetA.createRow(i);
                for (int c = 0; c < cells.length; c++) {
                    cell = row.createCell(c);
                    cell.setCellValue("小" + cells[c] + c);
                }
            }
        }
        File file = new File(Environment.getExternalStorageDirectory()
                + File.separator + "scdz" + File.separator + "newfile.xls");
        try {
            file.createNewFile();
            FileOutputStream outS = new FileOutputStream(file);
            hssfWorkbook.write(outS);
            outS.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
