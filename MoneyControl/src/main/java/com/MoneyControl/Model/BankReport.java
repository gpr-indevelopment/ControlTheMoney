package com.MoneyControl.Model;

import com.MoneyControl.Utils.Utils;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

@NoArgsConstructor
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

    public void setDate(LocalDate date) {
        this.date = date;
    }

    public String getStringDate(){
        return this.date.format(formatter);
    }

    public void setCredit(Double credit) {
        this.credit = credit;
    }

    public void setBalance(double balance) {
        this.balance = balance;
    }

    public void setDocumentNumber(Long documentNumber) {
        this.documentNumber = documentNumber;
    }

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

    public void setDebit(Double debit) {
        this.debit = debit;
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

    public BankReport clone(){
        BankReport cloneReport = new BankReport();
        cloneReport.setDate(this.date);
        cloneReport.setDebit(this.debit);
        cloneReport.setCredit(this.credit);
        cloneReport.setBalance(this.balance);
        cloneReport.setDescription(this.description);
        cloneReport.setSituation(this.situation);
        cloneReport.setDocumentNumber(this.documentNumber);
        return cloneReport;
    }
}
