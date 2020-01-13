package br.com.rodrigopereira.exportexcel.controlers;


import br.com.rodrigopereira.exportexcel.services.PlanilhaService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
@RequestMapping("/planilhas")
public class PlanilhaController {

    private final PlanilhaService planilhaService;

    @Autowired
    public PlanilhaController(PlanilhaService planilhaService) {
        this.planilhaService = planilhaService;
    }

    @GetMapping()
    public ResponseEntity<Void> gerarPlanilha() throws IOException {
        planilhaService.gerarPlanilha();
        return new ResponseEntity<>(HttpStatus.OK);
    }

}
