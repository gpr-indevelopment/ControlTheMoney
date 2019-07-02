package com.MoneyControl;

import lombok.Getter;
import lombok.Setter;

import java.util.Date;

public class BankReport {

    @Getter @Setter
    private Date date;

    @Getter @Setter
    private String description;

    @Getter @Setter
    private Long documentNumber;

    @Getter @Setter
    private String situation;

    @Getter @Setter
    private Integer credit;

    @Getter @Setter
    private Integer debit;

    @Getter @Setter
    private Integer balance;
}
