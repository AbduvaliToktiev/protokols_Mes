package com.example.protokols_mes.service;

import com.example.protokols_mes.dao.ProtokolRepository;
import com.example.protokols_mes.entity.Protokol;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class ProtokolService {

    @Autowired
    private ProtokolRepository protokolRepository;

    public void save(Protokol protokol) {
        protokol.setCompany("ОсОО <<МЭС>>");
        protokol.setDimensionMethod("Сопротивление изоляции обмоток относительно корпуса и между собой, измеренное");
        protokol.setTyp("M4100/5");
        protokol.setFactoryNumberIzol("№ 463232");
        protokol.setSoprotivlenieR15("Сопротивл-е R15Мом");
        protokol.setSoprotivlenieR60("Сопротивл-е R60Мом");
        protokol.setCooling("Естественное масляное");
        protokol.setCheckMethod("Увлажненность обомток проверена методом абсорбции. Измеренный кэф-т 3≥1");
        protokol.setConstantCurrency("Сопротивление обмоток постоянному току измерено Р4833");
        protokol.setFactoryNumberStatic("№ 14388");
        protokol.setChecked("Проверил");
        protokol.setPhaseA_B("A-B");
        protokol.setPhaseB_C("B-C");
        protokol.setPhaseC_A("C-A");
        this.protokolRepository.save(protokol);
    }

    public List<Protokol> findAllProtokol() {
        return this.protokolRepository.findAll();
    }

    public Protokol findById(Long protokolId) {
        return this.protokolRepository.findById(protokolId).orElse(null);
    }

    public List<Protokol> getProtokolByName(String keyword) {
        return this.protokolRepository.findByKeyword(keyword);
    }
}
