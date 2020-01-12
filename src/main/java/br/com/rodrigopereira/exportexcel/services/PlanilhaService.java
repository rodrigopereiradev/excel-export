package br.com.rodrigopereira.exportexcel.services;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

@Service
public class PlanilhaService {

    public void gerarPlanilha() throws IOException {

        Workbook planilha = new XSSFWorkbook();

        Sheet sheet = planilha.createSheet("Temas");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);

        Row cabecalho = sheet.createRow(0);

        CellStyle estiloNomesColunas = planilha.createCellStyle();
        estiloNomesColunas.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        estiloNomesColunas.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont fonte = ((XSSFWorkbook) planilha).createFont();
        fonte.setFontName("Arial");
        fonte.setFontHeightInPoints((short) 16);
        fonte.setBold(Boolean.TRUE);

        estiloNomesColunas.setFont(fonte);

        Cell nomesColunas = cabecalho.createCell(0);
        nomesColunas.setCellValue("Nome");
        nomesColunas.setCellStyle(estiloNomesColunas);

        nomesColunas = cabecalho.createCell(1);
        nomesColunas.setCellValue("Age");
        nomesColunas.setCellStyle(estiloNomesColunas);

        CellStyle estiloLinhas = planilha.createCellStyle();
        estiloLinhas.setWrapText(Boolean.TRUE);

        Row linhas = sheet.createRow(1);
        Cell celula = linhas.createCell(0);
        celula.setCellValue("John Smith");
        celula.setCellStyle(estiloLinhas);

        celula = linhas.createCell(1);
        celula.setCellValue(20);
        celula.setCellStyle(estiloLinhas);

        File diretorioAtual = new File(".");
        String path = diretorioAtual.getAbsolutePath();
        String localizacaoArquivo = path.substring(0, path.length() - 1) + "teste.xlsx";

        FileOutputStream outputStream = new FileOutputStream(localizacaoArquivo);
        planilha.write(outputStream);
        planilha.close();

    }


}
