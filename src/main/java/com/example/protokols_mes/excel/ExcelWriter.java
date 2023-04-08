package com.example.protokols_mes.excel;

import com.example.protokols_mes.entity.Protokol;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.Objects;

@Service
public class ExcelWriter {
    public void excelWriteProtokol(Protokol protokol) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("протокол тр-р " + protokol.getPowerKVA().intValue() + " " + protokol.getPower().charAt(0) + "" + protokol.getPower().charAt(1) + " " + protokol.getFactoryNumber());

        sheet.setColumnWidth(0, 1000);

        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 2));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 2));
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 2));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 2));

        Row row = sheet.createRow((short) 1);

        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setFontName("Cambria");

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);

        Cell cell = row.createCell(0);
        cell.setCellValue("Общество с ограниченной");
        cell.setCellStyle(cellStyle);

        Row row1 = sheet.createRow((short) 2);

        CellStyle cellStyle1 = workbook.createCellStyle();
        cellStyle1.setFont(font);

        cell = row1.createCell(0);
        cell.setCellValue("ответственностью");
        cell.setCellStyle(cellStyle1);

        Row row2 = sheet.createRow((short) 3);

        CellStyle cellStyle2 = workbook.createCellStyle();
        cellStyle2.setFont(font);

        cell = row2.createCell(0);
        cell.setCellValue("<<МЭС>> (Машэлектроснаб)");
        cell.setCellStyle(cellStyle2);

        Row row3 = sheet.createRow((short) 4);

        CellStyle cellStyle3 = workbook.createCellStyle();
        cellStyle3.setFont(font);

        cell = row3.createCell(0);
        cell.setCellValue("Пусконаладочный участок");
        cell.setCellStyle(cellStyle3);

        Row row4 = sheet.createRow((short) 5);

        CellStyle cellStyle4 = workbook.createCellStyle();
        CreationHelper creationHelper = workbook.getCreationHelper();
        cellStyle4.setDataFormat(creationHelper.createDataFormat().getFormat("<<dd>> MMMM yyyy"));
        cellStyle4.setAlignment(HorizontalAlignment.LEFT);
        cellStyle4.setFont(font);

        Cell cell1 = row4.createCell(0);
        cell1.setCellValue(protokol.getDate());
        cell1.setCellStyle(cellStyle4);

        sheet.setColumnWidth(8, 4000);

        Row row5 = sheet.getRow((short) 1);

        CellStyle cellStyle5 = workbook.createCellStyle();
        cellStyle5.setFont(font);

        cell = row5.createCell(8);
        cell.setCellValue("Предприятие");
        cell.setCellStyle(cellStyle5);

        Font font1 = workbook.createFont();
        font1.setFontHeightInPoints((short) 12);
        font1.setBold(true);
        font1.setFontName("Cambria");

        Row row6 = sheet.getRow((short) 2);

        CellStyle cellStyle6 = workbook.createCellStyle();
        cellStyle6.setFont(font1);

        cell = row6.createCell(8);
        cell.setCellValue("ОсОО <<МЭС>>");
        cell.setCellStyle(cellStyle6);

        Row row7 = sheet.getRow((short) 3);

        CellStyle cellStyle7 = workbook.createCellStyle();
        cellStyle7.setFont(font);

        cell = row7.createCell(8);
        cell.setCellValue("Объект : " + protokol.getObject());
        cell.setCellStyle(cellStyle7);

        sheet.addMergedRegion(new CellRangeAddress(7, 7, 0, 8));

        Row row8 = sheet.createRow((short) 7);

        CellStyle cellStyle8 = workbook.createCellStyle();
        cellStyle8.setAlignment(HorizontalAlignment.CENTER);
        cellStyle8.setFont(font1);

        Row row77 = sheet.createRow(6);

        Cell cell77 = row77.createCell(0);
        cell77.getRow().setHeightInPoints(cell77.getSheet().getDefaultRowHeightInPoints() * 3);

        Cell cell2 = row8.createCell(0);
        cell2.setCellValue(protokol.getName());
        cell2.setCellStyle(cellStyle8);

        sheet.addMergedRegion(new CellRangeAddress(8, 8, 0, 8));

        Row row9 = sheet.createRow((short) 8);

        CellStyle cellStyle9 = workbook.createCellStyle();
        cellStyle9.setAlignment(HorizontalAlignment.CENTER);
        cellStyle9.setFont(font1);

        cell2 = row9.createCell(0);
        cell2.setCellValue("Испытания силового маслонаполненного");
        cell2.setCellStyle(cellStyle9);

        sheet.addMergedRegion(new CellRangeAddress(9, 9, 0, 8));

        Row row10 = sheet.createRow((short) 9);

        CellStyle cellStyle10 = workbook.createCellStyle();
        cellStyle10.setAlignment(HorizontalAlignment.CENTER);
        cellStyle10.setFont(font1);

        cell2 = row10.createCell(0);
        cell2.setCellValue("трансформатора мощностью " + protokol.getPowerKVA().intValue() + " кВА");
        cell2.setCellStyle(cellStyle10);

        Row row11 = sheet.createRow((short) 11);

        CellStyle cellStyle11 = workbook.createCellStyle();
        cellStyle11.setFont(font);

        cell = row11.createCell(0);
        cell.setCellValue("1. Паспортные данные трансформатора");
        cell.setCellStyle(cellStyle11);

        Row row12 = sheet.createRow(13);

        CellStyle cellStyle12 = workbook.createCellStyle();
        cellStyle12.setWrapText(true);
        cellStyle12.setVerticalAlignment(VerticalAlignment.TOP);
        cellStyle12.setAlignment(HorizontalAlignment.LEFT);
        cellStyle12.setBorderRight(BorderStyle.valueOf((short) 1));
        cellStyle12.setBorderLeft(BorderStyle.valueOf((short) 1));
        cellStyle12.setBorderTop(BorderStyle.valueOf((short) 1));
        cellStyle12.setBorderBottom(BorderStyle.valueOf((short) 1));
        cellStyle12.setFont(font1);

        sheet.setColumnWidth(2, 4000);
        sheet.setColumnWidth(3, 3500);
        sheet.setColumnWidth(4, 4000);
        sheet.setColumnWidth(5, 3000);
        sheet.setColumnWidth(6, 3400);
        sheet.setColumnWidth(7, 3000);

        Cell cell3 = row12.createCell(1);
        cell3.getRow().setHeightInPoints(cell3.getSheet().getDefaultRowHeightInPoints() * 2);
        cell3.setCellValue("Тип");
        cell3.setCellStyle(cellStyle12);

        cell3 = row12.createCell(2);
        cell3.setCellValue("Зав.№");
        cell3.setCellStyle(cellStyle12);

        cell3 = row12.createCell(3);
        cell3.setCellValue("Мощность кВА");
        cell3.setCellStyle(cellStyle12);

        cell3 = row12.createCell(4);
        cell3.getRow().setHeightInPoints(cell3.getSheet().getDefaultRowHeightInPoints() * 2);
        cell3.setCellValue("Напряяж-е  кВ");
        cell3.setCellStyle(cellStyle12);

        cell3 = row12.createCell(5);
        cell3.setCellValue("Ток.            А");
        cell3.setCellStyle(cellStyle12);

        cell3 = row12.createCell(6);
        cell3.setCellValue("Ек.              %");
        cell3.setCellStyle(cellStyle12);

        cell3 = row12.createCell(7);
        cell3.setCellValue("Группа соед-я");
        cell3.setCellStyle(cellStyle12);

        cell3 = row12.createCell(8);
        cell3.setCellValue("Охлаждение");
        cell3.setCellStyle(cellStyle12);

        Row row13 = sheet.createRow(14);

        CellStyle cellStyle13 = workbook.createCellStyle();
        cellStyle13.setWrapText(true);
        cellStyle13.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle13.setAlignment(HorizontalAlignment.CENTER);
        cellStyle13.setBorderRight(BorderStyle.valueOf((short) 1));
        cellStyle13.setBorderLeft(BorderStyle.valueOf((short) 1));
        cellStyle13.setBorderTop(BorderStyle.valueOf((short) 1));
        cellStyle13.setBorderBottom(BorderStyle.valueOf((short) 1));
        cellStyle13.setFont(font);

        Cell cell4 = row13.createCell(1);
        cell4.getRow().setHeightInPoints(cell4.getSheet().getDefaultRowHeightInPoints() * 2);
        cell4.setCellValue(protokol.getTypeTransformator().getTranslate());
        cell4.setCellStyle(cellStyle13);

        cell4 = row13.createCell(2);
        cell4.setCellValue(protokol.getFactoryNumber());
        cell4.setCellStyle(cellStyle13);

        cell4 = row13.createCell(3);
        cell4.setCellValue(protokol.getPowerKVA().intValue());
        cell4.setCellStyle(cellStyle13);

        cell4 = row13.createCell(4);
        cell4.setCellValue(protokol.getPower());
        cell4.setCellStyle(cellStyle13);

        cell4 = row13.createCell(5);
        cell4.setCellValue(protokol.getTokA());
        cell4.setCellStyle(cellStyle13);

        cell4 = row13.createCell(6);
        cell4.setCellValue(protokol.getEk().doubleValue());
        cell4.setCellStyle(cellStyle13);

        cell4 = row13.createCell(7);
        cell4.setCellValue(protokol.getConnectionGroup());
        cell4.setCellStyle(cellStyle13);

        cell4 = row13.createCell(8);
        cell4.setCellValue(protokol.getCooling());
        cell4.setCellStyle(cellStyle13);

        Row row14 = sheet.createRow(16);

        CellStyle cellStyle14 = workbook.createCellStyle();
        cellStyle14.setFont(font);

        Cell cell5 = row14.createCell(0);
        cell5.setCellValue("2. " + protokol.getDimensionMethod());
        cell5.setCellStyle(cellStyle14);

        Row row15 = sheet.createRow(17);

        cell5 = row15.createCell(0);
        cell5.setCellValue("     мегомметром 2500 В " + " тип " + protokol.getTyp() + " зав." + protokol.getFactoryNumberIzol());
        cell5.setCellStyle(cellStyle14);

        Row row16 = sheet.createRow(19);

        cell5 = row16.createCell(0);
        cell5.setCellValue("Температура обмотки при измерений " + protokol.getTemperature() + "˚С");
        cell5.setCellStyle(cellStyle14);

        sheet.addMergedRegion(new CellRangeAddress(21, 21, 1, 2));
        sheet.addMergedRegion(new CellRangeAddress(21, 21, 3, 4));
        sheet.addMergedRegion(new CellRangeAddress(21, 21, 5, 6));
        sheet.addMergedRegion(new CellRangeAddress(21, 21, 7, 8));

        Row row17 = sheet.createRow(21);

        CellStyle cellStyle15 = workbook.createCellStyle();
        cellStyle15.setWrapText(true);
        cellStyle15.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle15.setAlignment(HorizontalAlignment.CENTER);
        cellStyle15.setBorderRight(BorderStyle.valueOf((short) 1));
        cellStyle15.setBorderLeft(BorderStyle.valueOf((short) 1));
        cellStyle15.setBorderTop(BorderStyle.valueOf((short) 1));
        cellStyle15.setBorderBottom(BorderStyle.valueOf((short) 1));
        cellStyle15.setFont(font1);

        Cell cell6 = row17.createCell(1);
        cell6.setCellValue("Схема измерения");
        cell6.setCellStyle(cellStyle15);

        cell6 = row17.createCell(2);
        cell6.setCellValue("Схема измерения");
        cell6.setCellStyle(cellStyle15);

        cell6 = row17.createCell(3);
        cell6.setCellValue("ВН-бак+НН");
        cell6.setCellStyle(cellStyle15);

        cell6 = row17.createCell(4);
        cell6.setCellValue("Схема измерения");
        cell6.setCellStyle(cellStyle15);

        cell6 = row17.createCell(5);
        cell6.setCellValue("НН-бак+ВН");
        cell6.setCellStyle(cellStyle15);

        cell6 = row17.createCell(6);
        cell6.setCellValue("Схема измерения");
        cell6.setCellStyle(cellStyle15);

        cell6 = row17.createCell(7);
        cell6.setCellValue("ВН-НН+бак");
        cell6.setCellStyle(cellStyle15);

        cell6 = row17.createCell(8);
        cell6.setCellValue("Схема измерения");
        cell6.setCellStyle(cellStyle15);

        sheet.addMergedRegion(new CellRangeAddress(22, 22, 1, 2));
        sheet.addMergedRegion(new CellRangeAddress(22, 22, 3, 4));
        sheet.addMergedRegion(new CellRangeAddress(22, 22, 5, 6));
        sheet.addMergedRegion(new CellRangeAddress(22, 22, 7, 8));

        Row row18 = sheet.createRow(22);

        cell6 = row18.createCell(1);
        cell6.setCellValue(protokol.getSoprotivlenieR15());
        cell6.setCellStyle(cellStyle15);

        cell6 = row18.createCell(2);
        cell6.setCellValue(protokol.getSoprotivlenieR15());
        cell6.setCellStyle(cellStyle15);

        CellStyle cellStyle16 = workbook.createCellStyle();
        cellStyle16.setWrapText(true);
        cellStyle16.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cellStyle16.setAlignment(HorizontalAlignment.CENTER);
        cellStyle16.setBorderRight(BorderStyle.valueOf((short) 1));
        cellStyle16.setBorderLeft(BorderStyle.valueOf((short) 1));
        cellStyle16.setBorderTop(BorderStyle.valueOf((short) 1));
        cellStyle16.setBorderBottom(BorderStyle.valueOf((short) 1));
        cellStyle16.setFont(font);

        Cell cell7 = row18.createCell(3);
        cell7.setCellValue(protokol.getVN_bak_NN_R15());
        cell7.setCellStyle(cellStyle16);

        cell7 = row18.createCell(4);
        cell7.setCellValue(protokol.getVN_bak_NN_R15());
        cell7.setCellStyle(cellStyle16);

        cell7 = row18.createCell(5);
        cell7.setCellValue(protokol.getNN_bak_VN_R15());
        cell7.setCellStyle(cellStyle16);

        cell7 = row18.createCell(6);
        cell7.setCellValue(protokol.getNN_bak_VN_R15());
        cell7.setCellStyle(cellStyle16);

        cell7 = row18.createCell(7);
        cell7.setCellValue(protokol.getVN_NN_bak_R15());
        cell7.setCellStyle(cellStyle16);

        cell7 = row18.createCell(8);
        cell7.setCellValue(protokol.getVN_NN_bak_R15());
        cell7.setCellStyle(cellStyle16);

        sheet.addMergedRegion(new CellRangeAddress(23, 23, 1, 2));
        sheet.addMergedRegion(new CellRangeAddress(23, 23, 3, 4));
        sheet.addMergedRegion(new CellRangeAddress(23, 23, 5, 6));
        sheet.addMergedRegion(new CellRangeAddress(23, 23, 7, 8));

        Row row19 = sheet.createRow(23);

        cell6 = row19.createCell(1);
        cell6.setCellValue(protokol.getSoprotivlenieR60());
        cell6.setCellStyle(cellStyle15);

        cell6 = row19.createCell(2);
        cell6.setCellValue(protokol.getSoprotivlenieR60());
        cell6.setCellStyle(cellStyle15);

        CellStyle cellStyle17 = workbook.createCellStyle();
        cellStyle17.setWrapText(true);
        cellStyle17.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cellStyle17.setAlignment(HorizontalAlignment.CENTER);
        cellStyle17.setBorderRight(BorderStyle.valueOf((short) 1));
        cellStyle17.setBorderLeft(BorderStyle.valueOf((short) 1));
        cellStyle17.setBorderTop(BorderStyle.valueOf((short) 1));
        cellStyle17.setBorderBottom(BorderStyle.valueOf((short) 1));
        cellStyle17.setFont(font);

        Cell cell8 = row19.createCell(3);
        cell8.setCellValue(protokol.getVN_bak_NN_R60());
        cell8.setCellStyle(cellStyle16);

        cell8 = row19.createCell(4);
        cell8.setCellValue(protokol.getVN_bak_NN_R60());
        cell8.setCellStyle(cellStyle16);

        cell8 = row19.createCell(5);
        cell8.setCellValue(protokol.getNN_bak_VN_R60());
        cell8.setCellStyle(cellStyle16);

        cell8 = row19.createCell(6);
        cell8.setCellValue(protokol.getNN_bak_VN_R60());
        cell8.setCellStyle(cellStyle16);

        cell8 = row19.createCell(7);
        cell8.setCellValue(protokol.getVN_NN_bak_R60());
        cell8.setCellStyle(cellStyle16);

        cell8 = row19.createCell(8);
        cell8.setCellValue(protokol.getVN_NN_bak_R60());
        cell8.setCellStyle(cellStyle16);

        Row row20 = sheet.createRow(25);

        cell5 = row20.createCell(0);
        cell5.setCellValue("3. " + protokol.getCheckMethod());
        cell5.setCellStyle(cellStyle14);

        Row row21 = sheet.createRow(26);

        cell5 = row21.createCell(0);
        cell5.setCellValue("4. " + protokol.getConstantCurrency() + " зав." + protokol.getFactoryNumberStatic());
        cell5.setCellStyle(cellStyle14);

        Row row22 = sheet.createRow(27);

        cell5 = row22.createCell(0);
        cell5.setCellValue("     при температуре " + protokol.getTemperatura() + "˚С");
        cell5.setCellStyle(cellStyle14);

        sheet.addMergedRegion(new CellRangeAddress(29, 29, 0, 1));
        sheet.addMergedRegion(new CellRangeAddress(29, 29, 2, 6));
        sheet.addMergedRegion(new CellRangeAddress(29, 30, 7, 8));
        sheet.addMergedRegion(new CellRangeAddress(30, 30, 0, 1));

        Row row23 = sheet.createRow(29);

        Cell cell9 = row23.createCell(0);
        cell9.setCellValue("Наимнов.обмоток");
        cell9.setCellStyle(cellStyle15);

        cell9 = row23.createCell(1);
        cell9.setCellValue("Наимнов.обмоток");
        cell9.setCellStyle(cellStyle15);

        cell9 = row23.createCell(2);
        cell9.setCellValue("Высогоко напряжения");
        cell9.setCellStyle(cellStyle15);

        cell9 = row23.createCell(3);
        cell9.setCellValue("Высогоко напряжения");
        cell9.setCellStyle(cellStyle15);

        cell9 = row23.createCell(4);
        cell9.setCellValue("Высогоко напряжения");
        cell9.setCellStyle(cellStyle15);

        cell9 = row23.createCell(5);
        cell9.setCellValue("Высогоко напряжения");
        cell9.setCellStyle(cellStyle15);

        cell9 = row23.createCell(6);
        cell9.setCellValue("Высогоко напряжения");
        cell9.setCellStyle(cellStyle15);

        cell9 = row23.createCell(7);
        cell9.setCellValue("Низкого           напряжения");
        cell9.setCellStyle(cellStyle15);

        cell9 = row23.createCell(8);
        cell9.setCellValue("Низкого           напряжения");
        cell9.setCellStyle(cellStyle15);

        Row row24 = sheet.createRow(30);

        cell9 = row24.createCell(0);
        cell9.setCellValue("Полож.перекл.");
        cell9.setCellStyle(cellStyle15);

        cell9 = row24.createCell(1);
        cell9.setCellValue("Полож.перекл.");
        cell9.setCellStyle(cellStyle15);

        cell9 = row24.createCell(2);
        cell9.setCellValue("I");
        cell9.setCellStyle(cellStyle15);

        cell9 = row24.createCell(3);
        cell9.setCellValue("II");
        cell9.setCellStyle(cellStyle15);

        cell9 = row24.createCell(4);
        cell9.setCellValue("III");
        cell9.setCellStyle(cellStyle15);

        cell9 = row24.createCell(5);
        cell9.setCellValue("IV");
        cell9.setCellStyle(cellStyle15);

        cell9 = row24.createCell(6);
        cell9.setCellValue("V");
        cell9.setCellStyle(cellStyle15);

        CellStyle cellStyle22 = workbook.createCellStyle();
        cellStyle22.setBorderLeft(BorderStyle.valueOf((short) 1));

        cell9 = row24.createCell(9);
        cell9.setCellValue("");
        cell9.setCellStyle(cellStyle22);

        sheet.addMergedRegion(new CellRangeAddress(31, 33, 0, 0));

        Row row25 = sheet.createRow(31);

        cell9 = row25.createCell(0);
        cell9.setCellValue("фазы");
        cell9.setCellStyle(cellStyle15);

        cell9 = row25.createCell(1);
        cell9.setCellValue(protokol.getPhaseA_B());
        cell9.setCellStyle(cellStyle15);

        cell9 = row25.createCell(2);
        cell9.setCellValue(protokol.getHighVoltageI_AB().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row25.createCell(3);
        cell9.setCellValue(protokol.getHighVoltageII_AB().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row25.createCell(4);
        cell9.setCellValue(protokol.getHighVoltageIII_AB().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row25.createCell(5);
        cell9.setCellValue(protokol.getHighVoltageIV_AB().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row25.createCell(6);
        cell9.setCellValue(protokol.getHighVoltageV_AB().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row25.createCell(7);
        cell9.setCellValue("а-в");
        cell9.setCellStyle(cellStyle15);

        cell9 = row25.createCell(8);
        cell9.setCellValue(protokol.getLowVoltage_AB());
        cell9.setCellStyle(cellStyle17);

        Row row26 = sheet.createRow(32);

        cell9 = row26.createCell(0);
        cell9.setCellValue("фазы");
        cell9.setCellStyle(cellStyle15);

        cell9 = row26.createCell(1);
        cell9.setCellValue(protokol.getPhaseB_C());
        cell9.setCellStyle(cellStyle15);

        cell9 = row26.createCell(2);
        cell9.setCellValue(protokol.getHighVoltageI_BC().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row26.createCell(3);
        cell9.setCellValue(protokol.getHighVoltageII_BC().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row26.createCell(4);
        cell9.setCellValue(protokol.getHighVoltageIII_BC().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row26.createCell(5);
        cell9.setCellValue(protokol.getHighVoltageIV_BC().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row26.createCell(6);
        cell9.setCellValue(protokol.getHighVoltageV_BC().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row26.createCell(7);
        cell9.setCellValue("в-с");
        cell9.setCellStyle(cellStyle15);

        cell9 = row26.createCell(8);
        cell9.setCellValue(protokol.getLowVoltage_BC());
        cell9.setCellStyle(cellStyle17);

        Row row27 = sheet.createRow(33);

        cell9 = row27.createCell(0);
        cell9.setCellValue("фазы");
        cell9.setCellStyle(cellStyle15);

        cell9 = row27.createCell(1);
        cell9.setCellValue(protokol.getPhaseC_A());
        cell9.setCellStyle(cellStyle15);

        cell9 = row27.createCell(2);
        cell9.setCellValue(protokol.getHighVoltageI_CA().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row27.createCell(3);
        cell9.setCellValue(protokol.getHighVoltageII_CA().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row27.createCell(4);
        cell9.setCellValue(protokol.getHighVoltageIII_CA().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row27.createCell(5);
        cell9.setCellValue(protokol.getHighVoltageIV_CA().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row27.createCell(6);
        cell9.setCellValue(protokol.getHighVoltageV_CA().doubleValue());
        cell9.setCellStyle(cellStyle17);

        cell9 = row27.createCell(7);
        cell9.setCellValue("с-а");
        cell9.setCellStyle(cellStyle15);

        cell9 = row27.createCell(8);
        cell9.setCellValue(protokol.getLowVoltage_CA());
        cell9.setCellStyle(cellStyle17);

        Row row28 = sheet.createRow(35);

        Cell cell10 = row28.createCell(0);

        CellStyle cellStyle18 = workbook.createCellStyle();
        cellStyle18.setVerticalAlignment(VerticalAlignment.TOP);
        cellStyle18.setFont(font1);

        if (!Objects.equals(protokol.getNote(), "")) {
            cell10 = row28.createCell(0);
            cell10.setCellValue("Примечание: " + protokol.getNote());
            cell10.setCellStyle(cellStyle18);
        } else {
            cell10.setCellValue("Примечание" + "________________________________________________________________________________________________________________________");
            cell10.setCellStyle(cellStyle18);
        }

        Row row29 = sheet.createRow(37);

        CellStyle cellStyle19 = workbook.createCellStyle();
        cellStyle19.setBorderTop(BorderStyle.valueOf((short) 1));
        cellStyle19.setBorderBottom(BorderStyle.valueOf((short) 1));
        cellStyle19.setFont(font);

        Cell cell11 = row29.createCell(0);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row29.createCell(1);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row29.createCell(2);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row29.createCell(3);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row29.createCell(4);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row29.createCell(5);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row29.createCell(6);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row29.createCell(7);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row29.createCell(8);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row29.createCell(9);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        sheet.setColumnWidth(1, 3000);

        sheet.addMergedRegion(new CellRangeAddress(38, 39, 0, 1));

        Row row30 = sheet.createRow(38);

        CellStyle cellStyle20 = workbook.createCellStyle();
        cellStyle20.setAlignment(HorizontalAlignment.CENTER);
        cellStyle20.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle20.setFont(font1);

        Cell cell12 = row30.createCell(0);
        cell12.setCellValue("Заключение");
        cell12.setCellStyle(cellStyle20);

        Row row31 = sheet.createRow(39);

        CellStyle cellStyle21 = workbook.createCellStyle();
        cellStyle21.setFont(font1);

        Cell cell13 = row31.createCell(3);
        cell13.setCellValue(protokol.getConclusion());
        cell13.setCellStyle(cellStyle21);

        Row row32 = sheet.createRow(40);

        cell11 = row32.createCell(0);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row32.createCell(1);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row32.createCell(2);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row32.createCell(3);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row32.createCell(4);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row32.createCell(5);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row32.createCell(6);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row32.createCell(7);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row32.createCell(8);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        cell11 = row32.createCell(9);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        Row row33 = sheet.createRow(42);

        cell12 = row33.createCell(3);
        cell12.setCellValue("Испытание и измерения произвели");
        cell12.setCellStyle(cellStyle20);

        Row row34 = sheet.createRow(45);

        cell12 = row34.createCell(1);
        cell12.setCellValue("Проверил");
        cell12.setCellStyle(cellStyle20);

        cell11 = row34.createCell(6);
        cell11.setCellValue(protokol.getFio());
        cell11.setCellStyle(cellStyle19);

        cell11 = row34.createCell(7);
        cell11.setCellValue("");
        cell11.setCellStyle(cellStyle19);

        InputStream inputStream = new FileInputStream("C:\\Users\\Admin\\OneDrive\\Документы\\Adobe\\лого ОсОО МЭС.png");

        byte[] bytes = IOUtils.toByteArray(inputStream);
        int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
        inputStream.close();

        CreationHelper creationHelper1 = workbook.getCreationHelper();

        Drawing<?> drawing = sheet.createDrawingPatriarch();
        // add a picture shape
        ClientAnchor anchor = creationHelper1.createClientAnchor();

        // set top-left corner of the picture,
        // subsequent call of Picture#resize() will operate relative to it
        anchor.setCol1(4);
        anchor.setRow1(0);
        Picture pict = drawing.createPicture(anchor, pictureIdx);

        // auto-size picture relative to its top-left corner
        pict.resize();

        try (FileOutputStream outputStream = new FileOutputStream("протокол тр-р " + protokol.getPowerKVA().intValue() + " " + protokol.getPower().charAt(0) + "" + protokol.getPower().charAt(1) + " " + protokol.getFactoryNumber() + ".xlsx")) {
            workbook.write(outputStream);
        }
    }
}
