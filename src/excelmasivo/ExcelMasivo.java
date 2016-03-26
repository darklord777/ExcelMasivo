/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelmasivo;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.GregorianCalendar;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 *
 * @author mario
 */
public class ExcelMasivo {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        String driver = "oracle.jdbc.OracleDriver";
        String user = "DRKL";
        String pass = "DRKL";
        String url = "jdbc:oracle:thin:@localhost:1521:XE";
        String query = "SELECT * FROM PRODUCTOS ORDER BY TO_NUMBER(SUBSTR(CODIGO_PRODUCTO,7))";
        Connection con;
        Statement st;
        ResultSet rs;
        ResultSetMetaData rsm;

        SXSSFWorkbook libro = new SXSSFWorkbook();
        SXSSFSheet hoja = libro.createSheet("Reporte");
        SXSSFRow fila;
        SXSSFCell celda;
        FileOutputStream out;
        int x = 0;

        CellStyle cs = libro.createCellStyle();
        cs.getFillForegroundColor();
        Font f = libro.createFont();
        //f.setBoldweight(Font.BOLDWEIGHT_BOLD);
        f.setFontHeightInPoints((short) 12);
        cs.setFont(f);

        try {
            Class.forName(driver);
            con = DriverManager.getConnection(url, user, pass);
            st = con.createStatement();
            rs = st.executeQuery(query);
            rsm = rs.getMetaData();
            while (rs.next()) {
                //crear la fila
                fila = hoja.createRow(x++);
                for (int i = 1; i <= rsm.getColumnCount(); i++) {
                    //recorrer las columnas
                    celda = fila.createCell(i);
                    celda.setCellStyle(cs);
                    celda.setCellValue(rs.getString(i));
                    //System.out.print(rs.getString(i));
                }
                //System.out.println();                
                if (x % 50000 == 0) {
                    System.out.println("Se procesaron:" + x);
                }
            }

            out = new FileOutputStream(new File("D:\\java\\Productos_" + GregorianCalendar.MILLISECOND + ".xlsx"));
            libro.write(out);
            out.close();
            System.out.println("Archivo generado con exito");
        } catch (ClassNotFoundException | SQLException | FileNotFoundException ex) {
            Logger.getLogger(ExcelMasivo.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExcelMasivo.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

}
