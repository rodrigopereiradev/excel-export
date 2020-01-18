package br.com.rodrigopereira.exportexcel.services;

import br.com.rodrigopereira.exportexcel.models.Pessoa;
import br.com.rodrigopereira.exportexcel.planilhas.PessoaPlanilha;
import br.com.rodrigopereira.exportexcel.planilhas.Planilha;
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

    public void gerarPlanilha() {

        List<Pessoa> pessoas = criaListaPessoas();
        Planilha planilha = new PessoaPlanilha(pessoas);
        Workbook planilhaGerada = planilha.gerarPlanilha();
        criarPlanilha(planilhaGerada);

    }

    private List<Pessoa> criaListaPessoas() {

        List<Pessoa> pessoas = new ArrayList<>();
        pessoas.add(new Pessoa("Rodrigo", 36));
        pessoas.add(new Pessoa("Leonardo", 26));
        pessoas.add(new Pessoa("Poliester", 35));
        pessoas.add(new Pessoa("Roberta", 19));

        return pessoas;
    }

    private void criarPlanilha(Workbook planilha) {
        try {
            File diretorioAtual = new File(".");
            String path = diretorioAtual.getAbsolutePath();
            String localizacaoArquivo = path.substring(0, path.length() - 1) + TEMAS_NOME_PLANILHA + EXTESAO_XLSX;

            FileOutputStream outputStream = new FileOutputStream(localizacaoArquivo);
            planilha.write(outputStream);
            planilha.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}
