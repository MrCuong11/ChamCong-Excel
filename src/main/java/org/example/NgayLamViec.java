package org.example;

import java.util.HashMap;
import java.util.Map;

public class NgayLamViec {
    private final int ngay;
    private final Map<String, Double> gioTheoCa = new HashMap<>();

    public NgayLamViec(int ngay) {
        this.ngay = ngay;
    }

    public void congGio(String ca, double gio) {
        gioTheoCa.put(ca, gioTheoCa.getOrDefault(ca, 0.0) + gio);
    }

    public double tongGio() {
        return gioTheoCa.values().stream().mapToDouble(Double::doubleValue).sum();
    }

    public double tinhTien(GiaCa giaCa) {
        return gioTheoCa.entrySet().stream()
                .mapToDouble(e -> e.getValue() * giaCa.getGia(e.getKey()))
                .sum();
    }

    public String toStringChiTiet(GiaCa giaCa) {
        return String.format("Ngày %d: Tổng giờ %.1f, tiền %.0f, chi tiết %s",
                ngay, tongGio(), tinhTien(giaCa), gioTheoCa);
    }
}