package org.example;


import java.util.HashMap;
import java.util.Map;

public class GiaCa {
    Map<String, Double> donGiaCa = new HashMap<>();

    public GiaCa(double cn, double gc, double tc, double gc1, double tc1, double wk) {
        donGiaCa.put("CN", cn);
        donGiaCa.put("GC", gc);
        donGiaCa.put("TC", tc);
        donGiaCa.put("GC1", gc1);
        donGiaCa.put("TC1", tc1);
        donGiaCa.put("WK-D", wk);
        donGiaCa.put("WK-N", wk);
    }

    public double getGia(String ca) {
        return donGiaCa.getOrDefault(ca, 0.0);
    }
}
