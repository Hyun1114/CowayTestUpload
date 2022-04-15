package com.excel.excel.service;

import java.io.IOException;
import java.util.List;

import com.excel.excel.helper.ExcelHelper;
import com.excel.excel.model.Coway;
import com.excel.excel.repository.CowayRepository;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ExcelService {

    @Autowired
    CowayRepository repository;
    public void save(MultipartFile file) {
        try {
            List<Coway> coways = ExcelHelper.excelToCoways(file.getInputStream());
            repository.saveAll(coways);
        } catch (IOException e) {
            throw new RuntimeException("fail to store excel data: " + e.getMessage());
        }
    }
    public List<Coway> getAllCoways() {
        return repository.findAll();
    }
}