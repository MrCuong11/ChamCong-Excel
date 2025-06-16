package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class ChamCongProcesser {

    // Cấu hình đơn giá theo từng nhân viên và từng loại ca
    static class GiaCa {
        Map<String, Double> donGiaCa = new HashMap<>();

        public GiaCa(double gc, double tc, double gc1, double tc1, double wk) {
            donGiaCa.put("GC", gc);
            donGiaCa.put("TC", tc);
            donGiaCa.put("GC1", gc1);
            donGiaCa.put("TC1", tc1);
            donGiaCa.put("WK-D", wk);
        }
    }

    // Dữ liệu mẫu bạn đã cung cấp
    static Map<String, GiaCa> bangGia = new HashMap<>();
    static {
        bangGia.put("Nguyen Van A", new GiaCa(216346, 180288, 156250, 240385, 240385));
        bangGia.put("Nguyen Van B", new GiaCa(72115, 108173, 93750, 144231, 144231));
        bangGia.put("Nguyen Van C", new GiaCa(43269, 64904, 56250, 86538, 86538));
        bangGia.put("Nguyen Van D", new GiaCa(38462, 57692, 50000, 76923, 76923));
    }

    public static void main(String[] args) throws Exception {
        FileInputStream fis = new FileInputStream(new File("C:\\Users\\NGUYEN MANH CUONG\\Downloads\\BangCong.xlsx"));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        // Dòng tiêu đề là dòng 4 (index 3), dữ liệu từ dòng 6 (index 5)
        for (int i = 5; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell cellName = row.getCell(2); // Cột Họ tên
            Cell cellLuong = row.getCell(15); // Cột Tổng lương

            if (cellName == null || cellLuong == null) continue;

            String hoTen = cellName.getStringCellValue().trim();
            if (hoTen.isEmpty()) continue;

            double tongLuongThucTe = 0.0;
            if (cellLuong.getCellType() == CellType.FORMULA) {
                if (cellLuong.getCachedFormulaResultType() == CellType.NUMERIC) {
                    tongLuongThucTe = cellLuong.getNumericCellValue();
                } else if (cellLuong.getCachedFormulaResultType() == CellType.STRING) {
                    String raw = cellLuong.getStringCellValue().replace(",", "").trim();
                    try {
                        tongLuongThucTe = Double.parseDouble(raw);
                    } catch (NumberFormatException e) {
                        System.err.println("Không đọc được lương cho nhân viên: " + hoTen);
                    }
                }
            } else if (cellLuong.getCellType() == CellType.NUMERIC) {
                tongLuongThucTe = cellLuong.getNumericCellValue();
            } else if (cellLuong.getCellType() == CellType.STRING) {
                String raw = cellLuong.getStringCellValue().replace(",", "").trim();
                try {
                    tongLuongThucTe = Double.parseDouble(raw);
                } catch (NumberFormatException e) {
                    System.err.println("Không đọc được lương cho nhân viên: " + hoTen);
                }
            }



            GiaCa giaCa = bangGia.getOrDefault(hoTen, new GiaCa(0, 0, 0, 0, 0));

            Map<String, Double> tongGioTheoCa = new HashMap<>();
            tongGioTheoCa.put("GC", 0.0);
            tongGioTheoCa.put("TC", 0.0);
            tongGioTheoCa.put("GC1", 0.0);
            tongGioTheoCa.put("TC1", 0.0);
            tongGioTheoCa.put("WK-D", 0.0);

            // Quét từ cột thứ 16 trở đi (ngày làm việc)
            for (int j = 16; j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                if (cell == null) continue;

                String val = "";
                if (cell.getCellType() == CellType.STRING) val = cell.getStringCellValue().trim();
                else if (cell.getCellType() == CellType.NUMERIC) val = String.valueOf(cell.getNumericCellValue());

                if (val.isEmpty()) continue;

                // Xử lý chuỗi có thể là "GC", "GC+GC", "WK-D", "4+4"
                String[] parts = val.split("\\+");
                for (String part : parts) {
                    part = part.trim().toUpperCase();
                    if (tongGioTheoCa.containsKey(part)) {
                        // Nếu là tên ca
                        tongGioTheoCa.put(part, tongGioTheoCa.get(part) + 4.0); // giả định mỗi ca là 4h
                    } else {
                        try {
                            double hours = Double.parseDouble(part);
                            tongGioTheoCa.put("GC", tongGioTheoCa.get("GC") + hours); // nếu chỉ là số, gán vào GC
                        } catch (NumberFormatException ignored) {}
                    }
                }
            }

            // Tính tổng tiền
            double tongLuongTinhToan = 0.0;
            for (String ca : tongGioTheoCa.keySet()) {
                double gio = tongGioTheoCa.get(ca);
                double gia = giaCa.donGiaCa.getOrDefault(ca, 0.0);
                tongLuongTinhToan += gio * gia;
            }

            System.out.println("===== " + hoTen + " =====");
            System.out.printf("Tổng lương file Excel: %.0f VND%n", tongLuongThucTe);
            System.out.printf("Tổng lương tính toán: %.0f VND%n", tongLuongTinhToan);
            System.out.printf("Chênh lệch: %.0f VND%n", tongLuongTinhToan - tongLuongThucTe);
            System.out.println("Chi tiết giờ theo ca: " + tongGioTheoCa);
            System.out.println();
        }

        workbook.close();
        fis.close();
    }
}
