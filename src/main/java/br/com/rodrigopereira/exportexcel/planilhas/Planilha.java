package br.com.rodrigopereira.exportexcel.planilhas;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public interface Planilha {

    Workbook gerarPlanilha();

    void gerarFolha();

    void gerarCabecalho();

    void gerarCabecalhoColunas();

    void gerarLinhas();

    default CellStyle gerarEstiloCabecalhoColunas(Workbook planilha) {
        CellStyle estiloNomesColunas = planilha.createCellStyle();

        estiloNomesColunas.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        estiloNomesColunas.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        estiloNomesColunas.setFont(gerarFonte(planilha, "Arial", 16, Boolean.TRUE));

        return estiloNomesColunas;
    }

    default CellStyle gerarEstiloCabecalho(Workbook planilha) {

        CellStyle estiloCabecalho = planilha.createCellStyle();
        estiloCabecalho.setFont(gerarFonte(planilha, "Arial", 16, Boolean.TRUE));

        return estiloCabecalho;
    }

    default XSSFFont gerarFonte(Workbook planilha, String nomeFonte, int tamanho, Boolean negrito) {
        XSSFFont fonte = ((XSSFWorkbook) planilha).createFont();
        fonte.setFontName(nomeFonte);
        fonte.setFontHeightInPoints((short) tamanho);
        fonte.setBold(negrito);
        return fonte;
    }

    default CellStyle gerarEstiloLinhas(Workbook planilha) {
        CellStyle estiloLinhas = planilha.createCellStyle();
        estiloLinhas.setWrapText(Boolean.TRUE);
        return estiloLinhas;
    }


    default void gerarCelula(Row linha, String valor, CellStyle estiloLinhas, int posicaoCelula) {
        Cell cell = linha.createCell(posicaoCelula);
        cell.setCellValue(valor);
        cell.setCellStyle(estiloLinhas);
    }

    default void gerarCelula(Row linha, Integer valor, CellStyle estiloLinhas, int posicaoCelula) {
        Cell cell = linha.createCell(posicaoCelula);
        cell.setCellValue(valor);
        cell.setCellStyle(estiloLinhas);
    }

}
