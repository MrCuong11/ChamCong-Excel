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

        double cn   = getDoubleCell(row.getCell(4));  // Cột E: CN
        double gc   = getDoubleCell(row.getCell(6));  // Cột G: GC
        double tc   = getDoubleCell(row.getCell(8));  // Cột I: TC
        double gc1  = getDoubleCell(row.getCell(10)); // Cột K: GC1
        double tc1  = getDoubleCell(row.getCell(12)); // Cột M: TC1
        double wk   = getDoubleCell(row.getCell(15)); // Cột P: WK

        GiaCa giaCa = new GiaCa(cn, gc, tc, gc1, tc1, wk);

        if (bangGia.containsKey(hoTen)) {
            System.err.println("Cảnh báo: Trùng tên nhân viên trong bảng giá: " + hoTen);
        }
        bangGia.put(hoTen, giaCa);
    }
}

    public static void main(String[] args) throws Exception {
        // Đọc file Excel
        FileInputStream fis = new FileInputStream(new File("C:/Users/NGUYEN MANH CUONG/Downloads/BangCong.xlsx"));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        docBangGia(sheet);

        // Duyệt qua từng dòng trong sheet (bắt đầu từ dòng 5)
        for (int i = 5; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell cellName = row.getCell(2);  // Cột C
            Cell cellLuong = row.getCell(16);  // Cột Q
            if (cellName == null || cellLuong == null) continue;

            String hoTen = cellName.getStringCellValue().trim();
            if (hoTen.isEmpty()) continue;

            double luongExcel = getDoubleCell(cellLuong);
            // GiaCa giaCa = bangGia.getOrDefault(hoTen, new GiaCa(0,0, 0, 0, 0, 0));
            GiaCa giaCa = bangGia.get(hoTen);
            if (giaCa == null) {
                System.err.println("Cảnh báo: Không tìm thấy bảng giá cho nhân viên: " + hoTen);
                giaCa = new GiaCa(0, 0, 0, 0, 0, 0);
            }
            NhanVien nv = new NhanVien(hoTen, luongExcel, giaCa);

            // Duyệt qua từng ngày trong tháng
            for (int ngay = 1; ngay <= 31; ngay++) {
                int colStart = getColumnStartIndex(ngay);
                if (colStart >= row.getLastCellNum()) break;

                // Tạo đối tượng NgayLamViec cho ngày đó
                NgayLamViec ngayLam = new NgayLamViec(ngay);
                boolean coLam = false;
                int soCot = getSoCotCuaNgay(ngay);

                // Duyệt qua từng cột trong 1 ngày
                for (int j = 0; j < soCot; j++) {
                    Cell cell = row.getCell(colStart + j);
                    String val = (cell == null) ? "" : getCellValue(cell);
                    if (!val.isEmpty()) {
                        coLam = true;
                        String ca = getCaName(ngay, j);
                        try {
                            double gio = Double.parseDouble(val);
                            ngayLam.congGio(ca, gio);
                        } catch (NumberFormatException ignored) {
                        }
                    }
                }
                if (coLam) nv.themNgayLam(ngayLam);
            }
            nv.inThongTin();
        }
        workbook.close();
        fis.close();
    }

    static int getSoCotCuaNgay(int ngay) {
        return switch (ngay) {
            case 5, 12, 19, 26 -> 2;
            case 1, 2, 3, 4, 6 -> 5;
            default -> 4;
        };
    }

    // Lấy vị trí cột bắt đầu của ngày đó
    static int getColumnStartIndex(int ngay) {
        int index = 17; // cột R
        for (int i = 1; i < ngay; i++) {
            index += getSoCotCuaNgay(i);
        }
        return index;
    }

    // Lấy tên ca từ vị trí cột
    static String getCaName(int ngay, int index) {
        return switch (getSoCotCuaNgay(ngay)) {
            case 5 -> switch (index) {
                case 0 -> "CN";
                case 1 -> "GC";
                case 2 -> "TC";
                case 3 -> "GC1";
                case 4 -> "TC1";
                default -> "GC";
            };
            case 2 -> (index == 0) ? "WK-D" : "WK-N";
            default -> switch (index) {
                case 0 -> "GC";
                case 1 -> "TC";
                case 2 -> "GC1";
                case 3 -> "TC1";
                default -> "GC";
            };
        };
    }

    // Lấy giá trị số từ ô Excel
    static double getDoubleCell(Cell cell) {
        try {
            if (cell.getCellType() == CellType.NUMERIC) return cell.getNumericCellValue();
            if (cell.getCellType() == CellType.STRING) return Double.parseDouble(cell.getStringCellValue().replace(",", ""));
            // Nếu ô là công thức, kiểm tra xem có phải là số không
            if (cell.getCellType() == CellType.FORMULA) {
                // Nếu công thức kết quả là số, trả về giá trị số
                if (cell.getCachedFormulaResultType() == CellType.NUMERIC) return cell.getNumericCellValue();
                // Nếu công thức kết quả là chuỗi, chuyển đổi thành số
                if (cell.getCachedFormulaResultType() == CellType.STRING)
                    return Double.parseDouble(cell.getStringCellValue().replace(",", ""));
            }
        } catch (Exception e) {
            System.err.println("Không đọc được ô dữ liệu: " + e.getMessage());
        }
        return 0.0;
    }

    // Lấy giá trị chuỗi từ ô Excel
    static String getCellValue(Cell cell) {
        if (cell.getCellType() == CellType.STRING) return cell.getStringCellValue().trim();
        if (cell.getCellType() == CellType.NUMERIC) return String.valueOf(cell.getNumericCellValue());
        return "";
    }
}