package com.excel.excel.model;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

import lombok.Data;

@Entity
@Table(name = "CowayTest2")
@Data
public class Coway {
    
    @Id
    @Column(name = "number")
    private long number;

    @Column(name = "testContent")
    private String testContent;

    @Column(name = "testType")
    private String testType;

    @Column(name = "groupColor")
    private String groupColor;

    @Column(name = "conUse")
    private String conUse;

    @Column(name = "blockName")
    private String blockName;

    @Column(name = "testApi")
    private String testApi;

    @Column(name = "testId")
    private String testId;

    @Column(name = "testTime")
    private String testTime;

    @Column(name = "ecMin")
    private String ecMin;

    @Column(name = "ecMax")
    private String ecMax;

    @Column(name = "vcMin")
    private String vcMin;

    @Column(name = "vcMax")
    private String vcMax;

    @Column(name = "ngMin")
    private String ngMin;

    @Column(name = "ngMax")
    private String ngMax;

    @Column(name = "vaMin")
    private String vaMin;

    @Column(name = "vaMax")
    private String vaMax;

    @Column(name = "vafix")
    private String vafix;

    @Column(name = "testInfo")
    private String testInfo;
    

}