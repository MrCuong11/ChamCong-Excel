package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class ChamCongProcessor {

    static Map<String, GiaCa> bangGia = new HashMap<>();

    static void docBangGia(Sheet sheet) {
        int dongGiaCa = 5; // giá ca
        int cotBatDau = 3; // tiêu đề bảng giá

        Row rowTieuDe = sheet.getRow(dongGiaCa);
        // lưu trữ ca
        List<String> danhSachCa = new ArrayList<>();
        // xác định cột $
        List<Integer> viTriCotGiaCa = new ArrayList<>();
        // lưu nhóm ca (WK-D, WK-N)
        Map<String, List<String>> nhomCa = new LinkedHashMap<>();

        //start
        int col = cotBatDau;
        while (col < rowTieuDe.getLastCellNum()) {

            String val = getCellValue(rowTieuDe.getCell(col));
//            nếu val chứa $
            if (val.equalsIgnoreCase("$")) {
                List<String> caCon = new ArrayList<>();
                int j = col - 1;
                while (j >= cotBatDau) {
                    String ca = getCellValue(rowTieuDe.getCell(j));
                    if (ca.equalsIgnoreCase("$")) break;
                    caCon.add(0, ca);
                    j--;
                }
                String caGop = (caCon.size() == 1) ? caCon.get(0) : "WK";
                // set nhom ca -> các ca con
                nhomCa.put(caGop, caCon);
                danhSachCa.add(caGop);
                viTriCotGiaCa.add(col);
            }
            col++;// next sang cột tiếp
        }

        //xet ca cho nhân viên để lưu bảng giá
        for (int i = dongGiaCa + 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell cellName = row.getCell(2); // Cột C
            if (cellName == null) continue;

            String hoTen = cellName.getStringCellValue().trim();
            if (hoTen.isEmpty()) continue;

            GiaCa giaCa = new GiaCa();
            for (int j = 0; j < danhSachCa.size(); j++) {
                int colIndex = viTriCotGiaCa.get(j);
                Cell cellGia = row.getCell(colIndex);
                double gia = getDoubleCell(cellGia);
                //set giá tiền cho ca
                giaCa.setGiaCa(danhSachCa.get(j), gia);
            }

            //xét ca cho nhân viên
            bangGia.put(hoTen, giaCa);
        }
    }

    public static void main(String[] args) throws Exception {
        FileInputStream fis = new FileInputStream(new File("C:/Users/NGUYEN MANH CUONG/Downloads/BangCong.xlsx"));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        docBangGia(sheet);

        Row rowNgay = sheet.getRow(3);
        Row rowCa = sheet.getRow(5);

        //duyệt ca trong ngày làm
        List<NgayVaCot> ngayVaCotList = new ArrayList<>();
        int colTongLuong = 16;
        int col = colTongLuong + 1; // cột bắt đầu sau cột tổng lương
        while (col <= rowNgay.getLastCellNum()) {
            Cell cell = rowNgay.getCell(col);
            // nếu là ngày (1,2,3,...)
            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                int ngay = (int) cell.getNumericCellValue();
                List<String> cacCa = new ArrayList<>();
                int nextCol = col;
                //tim các ca trong ngày, nếu thấy ngày mới thì dừng
                while (nextCol <= rowNgay.getLastCellNum()) {
                    Cell c = rowNgay.getCell(nextCol);
                    //next col != col => dừng (ngày mới)
                    if (c != null && c.getCellType() == CellType.NUMERIC && nextCol != col) break;
                    cacCa.add(getCellValue(rowCa.getCell(nextCol)));
                    nextCol++;
                }
                //tạo list object ngày và cột
                ngayVaCotList.add(new NgayVaCot(ngay, col, cacCa));
                col = nextCol;
            } else {
                col++;
            }
        }

        // duyệt qua nhân viên để tính công, giờ làm
        for (int i = 6; i <= sheet.getLastRowNum(); i++) {
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
                giaCa = new GiaCa();
            }
            NhanVien nv = new NhanVien(hoTen, luongExcel, giaCa);

            for (NgayVaCot ngayCot : ngayVaCotList) {
                NgayLamViec ngayLam = new NgayLamViec(ngayCot.ngay);
                boolean coLam = false;
                //duyệt ca trong ngày
                for (int j = 0; j < ngayCot.cacCa.size(); j++) {
                    int colIndex = ngayCot.startCol + j;
                    if (colIndex >= row.getLastCellNum()) break;
                    Cell cell = row.getCell(colIndex);
                    //get số giờ cột đó nếu ! null
                    String val = (cell == null) ? "" : getCellValue(cell);
                    if (!val.isEmpty()) {
                        try {
                            double gio = Double.parseDouble(val);
                            String caGoc = ngayCot.cacCa.get(j);
                            if (caGoc.equals("WK-D") || caGoc.equals("WK-N")) caGoc = "WK";
                            ngayLam.congGio(caGoc, gio);
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
            if (cell == null) return 0.0;
            //number -> number
            if (cell.getCellType() == CellType.NUMERIC) return cell.getNumericCellValue();
            // string -> string
            if (cell.getCellType() == CellType.STRING) return Double.parseDouble(cell.getStringCellValue().replace(",", ""));
            //coong thuc -> cong thuc
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
