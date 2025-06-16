package org.example;

import java.util.HashMap;
import java.util.Map;

public class NhanVienChamCong {
    public String hoTen;
    public double tongLuongThucTe;
    public Map<String, Double> tongGioTheoCa = new HashMap<>();
    public double tongLuongTinhToan;
    public double chenhLech;

    public NhanVienChamCong(String hoTen) {
        this.hoTen = hoTen;
        tongGioTheoCa.put("GC", 0.0);
        tongGioTheoCa.put("TC", 0.0);
        tongGioTheoCa.put("GC1", 0.0);
        tongGioTheoCa.put("TC1", 0.0);
        tongGioTheoCa.put("WK-D", 0.0);
    }

    public void tinhLuong(GiaCa giaCa) {
        tongLuongTinhToan = 0.0;
        for (String ca : tongGioTheoCa.keySet()) {
            double gio = tongGioTheoCa.get(ca);
            double gia = giaCa.donGiaCa.getOrDefault(ca, 0.0);
            tongLuongTinhToan += gio * gia;
        }
        chenhLech = tongLuongTinhToan - tongLuongThucTe;
    }

    public void inThongTin() {
        System.out.println("===== " + hoTen + " =====");
        System.out.printf("Tổng lương file Excel: %.0f VND%n", tongLuongThucTe);
        System.out.printf("Tổng lương tính toán: %.0f VND%n", tongLuongTinhToan);
        System.out.printf("Chênh lệch: %.0f VND%n", chenhLech);
        System.out.println("Chi tiết giờ theo ca: " + tongGioTheoCa);
        System.out.println();
    }
}
