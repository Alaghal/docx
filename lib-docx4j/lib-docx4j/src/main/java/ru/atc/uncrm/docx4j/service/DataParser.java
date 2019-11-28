package ru.atc.uncrm.docx4j.service;

import ru.atc.template.model.TemplateModel;

import java.util.List;
import java.util.Map;

public interface DataParser {

   List<Map<String,Object>> getListMapFromTemplateModel(String expression, TemplateModel model);

   Map<String,Object> getMapFromTemplateModel(String expression, TemplateModel model);


}
