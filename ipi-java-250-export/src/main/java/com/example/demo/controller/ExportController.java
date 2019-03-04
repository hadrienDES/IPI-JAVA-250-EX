package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.ClientService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.List;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;


    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        writer.println("Id;Nom;Prenom;Date de Naissance;Age");
        LocalDate now = LocalDate.now();
        for (Client client : allClients) {
            writer.println(
                    client.getId() + ";"
                            + client.getNom() + ";"
                            + client.getPrenom() + ";"
                            + client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")) + ";"
                            + (now.getYear() - client.getDateNaissance().getYear())
            );
        }
    }

    @GetMapping("/clients/xlsx")
    public void clientsXlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");

        Row headerRow = sheet.createRow(0);

        Cell cellHeaderId = headerRow.createCell(0);
        cellHeaderId.setCellValue("Id");

        Cell cellHeaderPrenom = headerRow.createCell(1);
        cellHeaderPrenom.setCellValue("Prénom");

        Cell cellHeaderNom = headerRow.createCell(2);
        cellHeaderNom.setCellValue("Nom");

        Cell cellHeaderDateNaissance = headerRow.createCell(3);
        cellHeaderDateNaissance.setCellValue("Date de naissance");

        int i = 1;
        for (Client client : allClients) {
            Row row = sheet.createRow(i);

            Cell cellId = row.createCell(0);
            cellId.setCellValue(client.getId());

            Cell cellPrenom = row.createCell(1);
            cellPrenom.setCellValue(client.getPrenom());

            Cell cellNom = row.createCell(2);
            cellNom.setCellValue(client.getNom());

            Cell cellDateNaissance = row.createCell(3);
            Date dateNaissance = Date.from(client.getDateNaissance().atStartOfDay(ZoneId.systemDefault()).toInstant());
            cellDateNaissance.setCellValue(dateNaissance);

            CellStyle cellStyleDate = workbook.createCellStyle();
            CreationHelper createHelper = workbook.getCreationHelper();
            cellStyleDate.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
            cellDateNaissance.setCellStyle(cellStyleDate);

            i++;
        }

        workbook.write(response.getOutputStream());
        workbook.close();

    }

    @GetMapping("/factures/xlsx")
    public void facturesXlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");
        List<Client> allClient = clientService.findAllClients();
        Workbook workbook = new XSSFWorkbook();

        int x = 0;

        for (Client client : allClient){
            Sheet clientSheet = workbook.createSheet("Client " + client.getId());

            Row clientHeaderRow = clientSheet.createRow(0);

            Cell cellHeaderIdClient = clientHeaderRow.createCell(0);
            cellHeaderIdClient.setCellValue("ID client");

            Cell cellHeaderPrenomClient = clientHeaderRow.createCell(1);
            cellHeaderPrenomClient.setCellValue("Prénom");

            Cell cellHeaderNomClient = clientHeaderRow.createCell(2);
            cellHeaderNomClient.setCellValue("Nom");

            Row clientRow = clientSheet.createRow(1);

            Cell cellIdClient = clientRow.createCell(0);
            cellIdClient.setCellValue(client.getId());

            Cell cellPrenomClient = clientRow.createCell(1);
            cellPrenomClient.setCellValue(client.getPrenom());

            Cell cellNomClient = clientRow.createCell(2);
            cellNomClient.setCellValue(client.getNom());

            for (Facture facture : client.getFactures()) {
                Double coutTotal = 0.0;

                Sheet sheet = workbook.createSheet("Facture " + facture.getId());

                Row headerRow = sheet.createRow(0);

                Cell cellHeaderArticle = headerRow.createCell(0);
                cellHeaderArticle.setCellValue("Article");

                Cell cellHeaderQuantite = headerRow.createCell(1);
                cellHeaderQuantite.setCellValue("Quantité");

                Cell cellHeaderPrixU = headerRow.createCell(2);
                cellHeaderPrixU.setCellValue("Prix unitaire");

                Cell cellHeaderTotal = headerRow.createCell(3);
                cellHeaderTotal.setCellValue("Total");

                int y = 1;

                for(LigneFacture ligneFacture : facture.getLigneFactures()){
                    Row row = sheet.createRow(y);

                    Cell cellHeaderArticle2 = row.createCell(0);
                    cellHeaderArticle2.setCellValue(ligneFacture.getArticle().getLibelle());

                    Cell cellHeaderQuantite2 = row.createCell(1);
                    cellHeaderQuantite2.setCellValue(ligneFacture.getQuantite());

                    Cell cellHeaderPrixU2 = row.createCell(2);
                    cellHeaderPrixU2.setCellValue(ligneFacture.getArticle().getPrix() + " €");

                    Cell cellHeaderTotal2 = row.createCell(3);
                    cellHeaderTotal2.setCellValue(String.format("%.2f", (ligneFacture.getQuantite()*ligneFacture.getArticle().getPrix())) + " €");

                    coutTotal += ligneFacture.getQuantite()*ligneFacture.getArticle().getPrix();
                    y++;
                }

                Row row = sheet.createRow(y);

                Cell cellTotal = row.createCell(0);
                cellTotal.setCellValue("Total " + String.format("%.2f", coutTotal) + "€");

                final Font font = sheet.getWorkbook().createFont ();
                font.setBold ( true );

                final CellStyle cellTotalStyle = workbook.createCellStyle();
                cellTotalStyle.setFont (font);

                cellTotalStyle.setFillBackgroundColor(IndexedColors.RED.getIndex());
                cellTotalStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                cellTotalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                cellTotalStyle.setBorderBottom(BorderStyle.THIN);
                cellTotalStyle.setBorderTop(BorderStyle.THIN);
                cellTotalStyle.setBorderLeft(BorderStyle.THIN);
                cellTotalStyle.setBorderRight(BorderStyle.THIN);
                cellTotal.setCellStyle(cellTotalStyle);
            }
        }

        workbook.write(response.getOutputStream());
        workbook.close();

    }
}

