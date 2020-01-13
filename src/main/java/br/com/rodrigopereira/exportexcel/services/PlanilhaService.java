package br.com.rodrigopereira.exportexcel.services;

import br.com.rodrigopereira.exportexcel.models.Pessoa;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

@Service
public class PlanilhaService {

    private static final String TEMAS_NOME_PLANILHA = "Pessoas";
    private static final String EXTESAO_XLSX = ".xlsx";

    private Workbook planilha;

    public void gerarPlanilha() {

        planilha = new XSSFWorkbook();
        List<Pessoa> pessoas = criaListaPessoas();

        Sheet folha = gerarFolha();
        Row cabecalho = folha.createRow(0);

        gerarCabecalhoColunas(cabecalho);
        criarLinhas(folha, pessoas);
        criarPlanilha();

    }

    private Sheet gerarFolha() {
        Sheet folha = planilha.createSheet(TEMAS_NOME_PLANILHA);
        folha.setColumnWidth(0, 6000);
        folha.setColumnWidth(1, 4000);
        return folha;
    }

    private void gerarCabecalhoColunas(Row cabecalho) {
        CellStyle estiloCabecalhoColunas = gerarEstiloCabecalhoColunas();

        gerarCelula(cabecalho, "Nome", estiloCabecalhoColunas, 0);
        gerarCelula(cabecalho, "Idade", estiloCabecalhoColunas, 1);

    }

    private CellStyle gerarEstiloLinhas() {
        CellStyle estiloLinhas = planilha.createCellStyle();
        estiloLinhas.setWrapText(Boolean.TRUE);
        return estiloLinhas;
    }

    private CellStyle gerarEstiloCabecalhoColunas() {
        CellStyle estiloNomesColunas = planilha.createCellStyle();

        estiloNomesColunas.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        estiloNomesColunas.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        estiloNomesColunas.setFont(gerarFonteCabecalhoColunas());

        return estiloNomesColunas;
    }

    private XSSFFont gerarFonteCabecalhoColunas() {
        XSSFFont fonte = ((XSSFWorkbook) planilha).createFont();
        fonte.setFontName("Arial");
        fonte.setFontHeightInPoints((short) 25);
        fonte.setBold(Boolean.TRUE);
        return fonte;
    }

    private void criarLinhas(Sheet folha, List<Pessoa> pessoas) {

        CellStyle estiloLinhas = gerarEstiloLinhas();

        int index = 1;

        for (Pessoa pessoa: pessoas) {

            Row linha = folha.createRow(index);

            gerarCelula(linha, pessoa.getNome(), estiloLinhas, 0);
            gerarCelula(linha, pessoa.getIdade(), estiloLinhas, 1);

            index++;
        }

    }

    private void criarPlanilha() {
        try {
//            File diretorioAtual = new File(".");
//            String path = diretorioAtual.getAbsolutePath();
            String path = "C:\\Users\\rodrigo.pereira\\desenvolvimento\\anexo_sfg_teste\\ ";
            String localizacaoArquivo = path.substring(0, path.length() - 1) + TEMAS_NOME_PLANILHA + EXTESAO_XLSX;

            FileOutputStream outputStream = new FileOutputStream(localizacaoArquivo);
            planilha.write(outputStream);
            planilha.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void gerarCelula(Row linha, String valor, CellStyle estiloLinhas, int posicaoCelula) {
            Cell cell = linha.createCell(posicaoCelula);
            cell.setCellValue(valor);
        cell.setCellStyle(estiloLinhas);
    }

    private void gerarCelula(Row linha, Integer valor, CellStyle estiloLinhas, int posicaoCelula) {
        Cell cell = linha.createCell(posicaoCelula);
        cell.setCellValue(valor);
        cell.setCellStyle(estiloLinhas);
    }

    private void gerarCelula(Row linha, Date valor, CellStyle estiloLinhas, int posicaoCelula) {
        Cell celula = linha.createCell(posicaoCelula);
        celula.setCellValue(valor);
        celula.setCellStyle(estiloLinhas);
    }


    private List<Pessoa> criaListaPessoas() {

        List<Pessoa> pessoas = new ArrayList<>();
        pessoas.add(new Pessoa("Rodrigo", 36));
        pessoas.add(new Pessoa("Leonardo", 26));
        pessoas.add(new Pessoa("Poliester", 35));
        pessoas.add(new Pessoa("Roberta", 19));

        return pessoas;
    }

}
