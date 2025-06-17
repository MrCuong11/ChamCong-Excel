package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class ChamCongProcessor {

    static Map<String, GiaCa> bangGia = new HashMap<>();

    static void docBangGia(Sheet sheet) {
        for (int i = 5; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell cellName = row.getCell(2); // Cột C: Họ tên (index 2)
            if (cellName == null) continue;
            String hoTen = cellName.getStringCellValue().trim();
            if (hoTen.isEmpty()) continue;

            double cn = getDoubleCell(row.getCell(4));  // Cột E: CN
            double gc = getDoubleCell(row.getCell(6));  // Cột G: GC
            double tc = getDoubleCell(row.getCell(8));  // Cột I: TC
            double gc1 = getDoubleCell(row.getCell(10)); // Cột K: GC1
            double tc1 = getDoubleCell(row.getCell(12)); // Cột M: TC1
            double wk = getDoubleCell(row.getCell(15)); // Cột P: WK

            GiaCa giaCa = new GiaCa(cn, gc, tc, gc1, tc1, wk);

            if (bangGia.containsKey(hoTen)) {
                System.err.println("Cảnh báo: Trùng tên nhân viên trong bảng giá: " + hoTen);
            }
            bangGia.put(hoTen, giaCa);
        }
    }

    public static void main(String[] args) throws Exception {
        FileInputStream fis = new FileInputStream(new File("C:/Users/NGUYEN MANH CUONG/Downloads/BangCong.xlsx"));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        docBangGia(sheet);

        // Dòng chứa số ngày (dòng 4 - index 3) và tên ca (dòng 6 - index 5)
        Row rowNgay = sheet.getRow(3);
        Row rowCa = sheet.getRow(5);

        // Xây dựng danh sách ngày và các ca tương ứng
        List<NgayVaCot> ngayVaCotList = new ArrayList<>();
        int col = 17; // cột R
        while (col <= rowNgay.getLastCellNum()) {
            Cell cell = rowNgay.getCell(col);
            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                int ngay = (int) cell.getNumericCellValue();
                List<String> cacCa = new ArrayList<>();
                int nextCol = col + 1;
                cacCa.add(getCellValue(rowCa.getCell(col)));
                while (nextCol <= rowNgay.getLastCellNum()) {
                    Cell c = rowNgay.getCell(nextCol);
                    if (c != null && c.getCellType() == CellType.NUMERIC) break;
                    cacCa.add(getCellValue(rowCa.getCell(nextCol)));
                    nextCol++;
                }
                ngayVaCotList.add(new NgayVaCot(ngay, col, cacCa));
                col = nextCol;
            } else {
                col++;
            }
        }

        for (int i = 5; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell cellName = row.getCell(2);
            Cell cellLuong = row.getCell(16);
            if (cellName == null || cellLuong == null) continue;

            String hoTen = cellName.getStringCellValue().trim();
            if (hoTen.isEmpty()) continue;

            double luongExcel = getDoubleCell(cellLuong);
            GiaCa giaCa = bangGia.get(hoTen);
            if (giaCa == null) {
                System.err.println("Cảnh báo: Không tìm thấy bảng giá cho nhân viên: " + hoTen);
                giaCa = new GiaCa(0, 0, 0, 0, 0, 0);
            }
            NhanVien nv = new NhanVien(hoTen, luongExcel, giaCa);

            // Duyệt qua từng ngày và xử lý từng ca
            for (NgayVaCot ngayCot : ngayVaCotList) {
                NgayLamViec ngayLam = new NgayLamViec(ngayCot.ngay);
                boolean coLam = false;
                for (int j = 0; j < ngayCot.cacCa.size(); j++) {
                    int colIndex = ngayCot.startCol + j;
                    if (colIndex >= row.getLastCellNum()) break;
                    Cell cell = row.getCell(colIndex);
                    String val = (cell == null) ? "" : getCellValue(cell);
                    if (!val.isEmpty()) {
                        try {
                            double gio = Double.parseDouble(val);
                            ngayLam.congGio(ngayCot.cacCa.get(j), gio);
                            coLam = true;
                        } catch (NumberFormatException ignored) {}
                    }
                }
                if (coLam) nv.themNgayLam(ngayLam);
            }
            nv.inThongTin();
        }
        workbook.close();
        fis.close();
    }


    static double getDoubleCell(Cell cell) {
        try {
            if (cell.getCellType() == CellType.NUMERIC) return cell.getNumericCellValue();
            if (cell.getCellType() == CellType.STRING) return Double.parseDouble(cell.getStringCellValue().replace(",", ""));
            if (cell.getCellType() == CellType.FORMULA) {
                if (cell.getCachedFormulaResultType() == CellType.NUMERIC) return cell.getNumericCellValue();
                if (cell.getCachedFormulaResultType() == CellType.STRING)
                    return Double.parseDouble(cell.getStringCellValue().replace(",", ""));
            }
        } catch (Exception e) {
            System.err.println("Không đọc được ô dữ liệu: " + e.getMessage());
        }
        return 0.0;
    }

    static String getCellValue(Cell cell) {
        if (cell == null) return "";
        if (cell.getCellType() == CellType.STRING) return cell.getStringCellValue().trim();
        if (cell.getCellType() == CellType.NUMERIC) return String.valueOf(cell.getNumericCellValue());
        return "";
    }
}
