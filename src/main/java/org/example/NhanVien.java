package org.example;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class NhanVien {
    private final String hoTen;
    private final double luongThucTe;
    private final GiaCa giaCa;
    private final List<NgayLamViec> dsNgayLam = new ArrayList<>();

    public NhanVien(String hoTen, double luongThucTe, GiaCa giaCa) {
        this.hoTen = hoTen;
        this.luongThucTe = luongThucTe;
        this.giaCa = giaCa;
    }

    public void themNgayLam(NgayLamViec ngay) {
        dsNgayLam.add(ngay);
    }

    public double tinhTongLuong() {
        return dsNgayLam.stream().mapToDouble(ngay -> ngay.tinhTien(giaCa)).sum();
    }

    public void inThongTin() {
        System.out.println("===== " + hoTen + " =====");
        System.out.printf("Tổng lương file Excel: %.0f VND%n", luongThucTe);
        System.out.printf("Tổng lương đã tính toán: %.0f VND%n", tinhTongLuong());
        System.out.printf("Chênh lệch: %.0f VND%n", tinhTongLuong() - luongThucTe);
        for (NgayLamViec ngay : dsNgayLam) {
            System.out.println(ngay.toStringChiTiet(giaCa));
        }
        System.out.println();
    }
}