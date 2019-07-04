package com.MoneyControl;

import lombok.Getter;
import lombok.Setter;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class BankReport {

    private DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");

    @Getter
    @Setter
    private LocalDate date;

    @Getter
    @Setter
    private String description;

    @Getter
    @Setter
    private Long documentNumber;

    @Getter
    @Setter
    private String situation;

    @Getter
    @Setter
    private Double credit;

    @Getter
    @Setter
    private Double debit;

    @Getter
    @Setter
    private double balance;

    public void setDate(String date) {
        this.date = LocalDate.parse(date, formatter);
    }

    public void setDocumentNumber(String documentNumber) {
        this.documentNumber = Long.parseLong(documentNumber);
    }

    public void setCredit(String credit) {
        if (credit.trim().isEmpty()) {
            this.credit = 0D;
        } else {
            this.credit = Double.parseDouble(credit);
        }
    }

    public void setDebit(String debit) {
        if (debit.trim().isEmpty()) {
            this.debit = 0D;
        } else {
            this.debit = Double.parseDouble(debit);
        }
    }

    public void setBalance(String balance) {
        if (balance.trim().isEmpty()) {
            this.balance = 0;
        } else {
            this.balance = Double.parseDouble(balance);
        }
    }

    public BankReport(String[] reportLine) {
        this.setDate(reportLine[0]);
        this.setDescription(Utils.trimDescription(reportLine[1]));
        this.setDocumentNumber(reportLine[2]);
        this.setSituation(reportLine[3]);
        this.setCredit(reportLine[4]);
        this.setDebit(reportLine[5]);
        this.setBalance(reportLine[6]);
    }

    public boolean isDebit(){
        boolean isDebit = false;
        if(this.debit != null && this.debit != 0)
        {
            isDebit = true;
        }
        return isDebit;
    }
}
