package com.example.demo.entity;

import javax.persistence.*;
import java.util.Set;

/**
 * Created by Alexandre on 09/04/2018.
 */
@Entity
public class Facture {

    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long id;

    @ManyToOne
    private Client client;

    @OneToMany(mappedBy = "facture")
    private Set<LigneFacture> ligneFactures;

    public double getTotal() {
        double prixTotal = 0;
        for (LigneFacture ligne : ligneFactures) {
            double prixLigne =
                    ligne.getArticle().getPrix()
                            * ligne.getQuantite();
            prixTotal += prixLigne;
        }
        return prixTotal;
    }

    public Long getId() {
        return id;
    }

    public Client getClient() {
        return client;
    }

    public void setClient(Client client) {
        this.client = client;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public Set<LigneFacture> getLigneFactures() {
        return ligneFactures;
    }

    public void setLigneFactures(Set<LigneFacture> ligneFactures) {
        this.ligneFactures = ligneFactures;
    }


}
