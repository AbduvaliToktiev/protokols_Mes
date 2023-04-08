package com.example.protokols_mes.entity;

import com.example.protokols_mes.enums.TransformatorType;
import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.springframework.format.annotation.DateTimeFormat;

import java.math.BigDecimal;
import java.util.Date;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Entity
@Table(name = Protokol.TABLE_NAME)
public class Protokol {

    public static final String TABLE_NAME = "PROTOKOLS";
    public static final String SEQ_NAME = TABLE_NAME + "_SEQ";

    @Id
    @SequenceGenerator(name = SEQ_NAME, sequenceName = SEQ_NAME, allocationSize = 1)
    @GeneratedValue(generator = SEQ_NAME)
    private Long id;

    @Column(name = "name")
    private String name;

    @Column(name = "company")
    private String company;

    @Column(name = "object")
    private String object;

    @Column(name = "date")
    @DateTimeFormat(pattern ="yyyy-MM-dd")
    private Date date;

    @Enumerated(EnumType.STRING)
    @Column(name = "type")
    private TransformatorType typeTransformator;

    @Column(name = "factory_number")
    private Integer factoryNumber;

    @Column(name = "power_kVA")
    private BigDecimal powerKVA;

    @Column(name = "power_kv")
    private String power;

    @Column(name = "tok_A")
    private String tokA;

    @Column(name = "Ek")
    private BigDecimal ek;

    @Column(name = "connection_group")
    private String connectionGroup;

    @Column(name = "cooling")
    private String cooling;

    @Column(name = "dimension_method")
    private String dimensionMethod;

    @Column(name = "typ")
    private String typ;

    @Column(name = "factory_number_izol")
    private String factoryNumberIzol;

    @Column(name = "temperature")
    private String temperature;

    @Column(name = "soprotivlenie_r15")
    private String soprotivlenieR15;

    @Column(name = "VN_bak_NN_R15")
    private Integer VN_bak_NN_R15;

    @Column(name = "NN_bak_VN_R15")
    private Integer NN_bak_VN_R15;

    @Column(name = "VN_NN_bak_R15")
    private Integer VN_NN_bak_R15;

    @Column(name = "soprotivlenie_r60")
    private String soprotivlenieR60;

    @Column(name = "VN_bak_NN_R60")
    private Integer VN_bak_NN_R60;

    @Column(name = "NN_bak_VN_R60")
    private Integer NN_bak_VN_R60;

    @Column(name = "VN_NN_bak_R60")
    private Integer VN_NN_bak_R60;

    @Column(name = "check_method")
    private String checkMethod;

    @Column(name = "constant_currency")
    private String constantCurrency;

    @Column(name = "factory_number_static")
    private String factoryNumberStatic;

    @Column(name = "temperatura")
    private String temperatura;

    @Column(name = "phase_a_b")
    private String phaseA_B;

    @Column(name = "high_voltage_I_a_b")
    private BigDecimal highVoltageI_AB;

    @Column(name = "high_voltage_II_a_b")
    private BigDecimal highVoltageII_AB;

    @Column(name = "high_voltage_III_a_b")
    private BigDecimal highVoltageIII_AB;

    @Column(name = "high_voltage_IV_a_b")
    private BigDecimal highVoltageIV_AB;

    @Column(name = "high_voltage_V_a_b")
    private BigDecimal highVoltageV_AB;

    @Column(name = "low_voltage_a_b")
    private Double lowVoltage_AB;

    @Column(name = "phase_b_c")
    private String phaseB_C;

    @Column(name = "high_voltage_I_b_c")
    private BigDecimal highVoltageI_BC;

    @Column(name = "high_voltage_II_b_c")
    private BigDecimal highVoltageII_BC;

    @Column(name = "high_voltage_III_b_c")
    private BigDecimal highVoltageIII_BC;

    @Column(name = "high_voltage_IV_b_c")
    private BigDecimal highVoltageIV_BC;

    @Column(name = "high_voltage_V_b_c")
    private BigDecimal highVoltageV_BC;

    @Column(name = "low_voltage_b_c")
    private Double lowVoltage_BC;

    @Column(name = "phase_c-a")
    private String phaseC_A;

    @Column(name = "high_voltage_I_c_a")
    private BigDecimal highVoltageI_CA;

    @Column(name = "high_voltage_II_c_a")
    private BigDecimal highVoltageII_CA;

    @Column(name = "high_voltage_III_c_a")
    private BigDecimal highVoltageIII_CA;

    @Column(name = "high_voltage_IV_c_a")
    private BigDecimal highVoltageIV_CA;

    @Column(name = "high_voltage_V_c_a")
    private BigDecimal highVoltageV_CA;

    @Column(name = "low_voltage_c_a")
    private Double lowVoltage_CA;

    @Column(name = "note")
    private String note;

    @Column(name = "conclusion")
    private String conclusion;

    @Column(name = "checked")
    private String checked;

    @Column(name = "fio")
    private String fio;
}
