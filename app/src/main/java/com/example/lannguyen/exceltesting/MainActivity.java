package com.example.lannguyen.exceltesting;

import android.Manifest;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.os.Environment;
import android.support.annotation.NonNull;
import android.support.v4.app.ActivityCompat;
import android.support.v7.app.AppCompatActivity;
import android.util.Log;
import android.view.View;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class MainActivity extends AppCompatActivity {

    public static final String PATH_FOLDER_ROOT = Environment.getExternalStorageDirectory() + File.separator + "PosTablet" + File.separator;
    public static final String PATH_FOLDER_DOWNLOAD = PATH_FOLDER_ROOT + "Download" + File.separator;
    public static final String PATH_FOLDER_BAO_CAO_CHI_TIET = PATH_FOLDER_ROOT + "BaoCaoChiTiet" + File.separator;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        TextView textView = findViewById(R.id.textCheck);
        textView.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                if (ActivityCompat.checkSelfPermission(MainActivity.this,
                        Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
                    ActivityCompat.requestPermissions(MainActivity.this,
                            new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE}, 1);
                } else {
                    saveExcelFile();
                }
            }
        });
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        if (requestCode == 1 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
            saveExcelFile();
        }
    }

    void saveExcelFile() {

        Date date = new Date();
        String strDateFormat = "dd_MM_yyyy";
        SimpleDateFormat sdf = new SimpleDateFormat(strDateFormat);
        String fileName = "Bao_cao_thu_chi_tiet_" + sdf.format(date) + ".xls";
        boolean success = false;
        //New Workbook
        Workbook wb = new HSSFWorkbook();

        Cell c = null;

        //Cell style for header row
        CellStyle cs = wb.createCellStyle();
        cs.setFillForegroundColor(HSSFColor.LIME.index);
        cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        CellStyle cs1 = wb.createCellStyle();
        cs1.setFillForegroundColor(HSSFColor.WHITE.index);
        cs1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        //New Sheet
        Sheet sheet1 = null;
        sheet1 = wb.createSheet("myOrder");

        // Generate column headings
        Row row1 = sheet1.createRow(0);
        c = row1.createCell(2);
        c.setCellValue("BÁO CÁO THU CHI TIẾT");
        c.setCellStyle(cs);

        Row row = sheet1.createRow(3);

        c = row.createCell(0);
        c.setCellValue("Mã Khách Hàng");
        c.setCellStyle(cs);

        c = row.createCell(1);
        c.setCellValue("Tên");
        c.setCellStyle(cs);

        c = row.createCell(2);
        c.setCellValue("Kỳ");
        c.setCellStyle(cs);

        c = row.createCell(3);
        c.setCellValue("Số Tiền");
        c.setCellStyle(cs);

        c = row.createCell(4);
        c.setCellValue("Ngày Thu");
        c.setCellStyle(cs);
        Bangexcel(c, cs1, sheet1);
        sheet1.setColumnWidth(0, (15 * 300));
        sheet1.setColumnWidth(1, (15 * 500));
        sheet1.setColumnWidth(2, (15 * 300));
        sheet1.setColumnWidth(3, (15 * 500));
        sheet1.setColumnWidth(4, (15 * 500));

        // Create a path where we will place our List of objects on external storage
        File directory = new File(PATH_FOLDER_ROOT);
        if (!directory.exists()) {
            directory.mkdirs();
        }else{
            Log.e("ch","1");
        }
        File file1 = new File(PATH_FOLDER_BAO_CAO_CHI_TIET);
        if (!file1.exists()) {
            file1.mkdirs();
        }else{
            Log.e("ch","2");
        }
        File file = new File(PATH_FOLDER_BAO_CAO_CHI_TIET, fileName);
        if (!file.exists()) {
            try {
                Log.e("ch","3");
                file.createNewFile();
            } catch (IOException e) {
                Log.e("ch","4");
                e.printStackTrace();
            }
        }
        FileOutputStream os = null;

        try {
            os = new FileOutputStream(file);
            wb.write(os);
            success = true;
            Toast.makeText(this, "Đã xuất ra file excel thành công", Toast.LENGTH_LONG).show();
        } catch (IOException e) {
            Log.w("FileUtils", "Error writing " + file, e);
        } catch (Exception e) {
            Log.w("FileUtils", "Failed to save file", e);
        } finally {
            try {
                if (null != os)
                    os.close();
            } catch (Exception ex) {
                Log.w("exception", "" + ex.getMessage());
            }
        }
    }

    private void Bangexcel(Cell c, CellStyle cs, Sheet sheet1) {

        Row row = sheet1.createRow(4);

        c = row.createCell(0);
        c.setCellValue(1);
        c.setCellStyle(cs);

        c = row.createCell(1);
        c.setCellValue(1);
        c.setCellStyle(cs);

        c = row.createCell(2);
        c.setCellValue(1);
        c.setCellStyle(cs);

        c = row.createCell(3);
        c.setCellValue(1);
        c.setCellStyle(cs);

        c = row.createCell(4);
        c.setCellValue(1);
        c.setCellStyle(cs);
    }
}

