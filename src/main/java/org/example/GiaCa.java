package org.example;


import java.util.HashMap;
import java.util.Map;

public class GiaCa {
    Map<String, Double> donGiaCa = new HashMap<>();

    public GiaCa() {}

    public void setGiaCa(String ca, double gia) {
        donGiaCa.put(ca, gia);
    }

    public double getGia(String ca) {
        return donGiaCa.getOrDefault(ca, 0.0);
    }

//    public Map<String, Double> getAllGia() {
//        return donGiaCa;
//    }
}