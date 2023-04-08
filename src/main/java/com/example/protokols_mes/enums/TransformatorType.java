package com.example.protokols_mes.enums;

import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter
@AllArgsConstructor
public enum TransformatorType {
    TM("ТМ"),
    TMG("ТМГ"),
    TMZ("ТМЗ"),
    TSL("ТСЛ");

    private String translate;
}

