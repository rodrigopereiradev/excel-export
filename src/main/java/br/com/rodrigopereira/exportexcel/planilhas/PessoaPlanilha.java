package br.com.rodrigopereira.exportexcel.planilhas;

import br.com.rodrigopereira.exportexcel.models.Pessoa;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

public class PessoaPlanilha implements Planilha {

    private final List<Pessoa> pessoas;
    private Workbook planilha;
    private Sheet folha;

    public PessoaPlanilha(List<Pessoa> pessoas) {
        this.pessoas = pessoas;
        this.planilha = new XSSFWorkbook();
    }

    @Override
    public Workbook gerarPlanilha() {

        gerarFolha();
        gerarCabecalho();
        gerarCabecalhoColunas();
        gerarLinhas();

        return planilha;
    }

    @Override
    public void gerarFolha() {

       folha = planilha.createSheet("Pessoas");
       folha.setDefaultColumnWidth(10);

    }

    @Override
    public void gerarCabecalho() {
        Row cabecalho = folha.createRow(0);
        CellStyle estilhoCabecalho = gerarEstiloCabecalho(planilha);
        gerarCelula(cabecalho, "Pessoas", estilhoCabecalho, 0);

        CellRangeAddress mesclargem = new CellRangeAddress(0, 0, 0, 1);

        folha.addMergedRegion(mesclargem);
    }

    @Override
    public void gerarCabecalhoColunas() {
        Row cabecalhoColunas = folha.createRow(1);

        CellStyle estiloCabecalhoColunas = gerarEstiloCabecalhoColunas(planilha);

        gerarCelula(cabecalhoColunas, "Nome", estiloCabecalhoColunas, 0);
        gerarCelula(cabecalhoColunas, "Idade", estiloCabecalhoColunas, 1);

    }

    @Override
    public void gerarLinhas() {
        CellStyle estiloLinhas = gerarEstiloLinhas(planilha);

        int index = 2;

        for (Pessoa pessoa: pessoas) {

            Row linha = folha.createRow(index);

            gerarCelula(linha, pessoa.getNome(), estiloLinhas, 0);
            gerarCelula(linha, pessoa.getIdade(), estiloLinhas, 1);

            index++;
        }
    }
}
