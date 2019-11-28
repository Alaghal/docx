package ru.atc.uncrm.docx4j.service.impl;

import ru.atc.template.model.TemplateModel;
import ru.atc.template.model.lazy.LazyMap;
import ru.atc.uncrm.docx4j.service.DataParser;


import java.util.*;
import java.util.stream.Collectors;


public class DataParserImpl implements DataParser {

    @Override
    public List<Map<String, Object>> getListMapFromTemplateModel(String expression, TemplateModel model) {

        List<Map<String, Object>> listResult = new ArrayList<>();
        Object obj = model.getCollection(expression);

        List listMap = (List) obj;
        for (Object lazymap :listMap) {
            Map<String, Object> mapResult = new HashMap<>();
            Map map =(LazyMap) lazymap;
            List<String> keyList = (List<String>) map.keySet().stream().collect(Collectors.toList());
            for (String key: keyList) {
                 Object value = map.get(key);
                 if(value != null){
                     mapResult.put(key, value);
                 }else {
                     mapResult.put(key,"");
                 }

            }

            listResult.add(mapResult);

        }

        return listResult;
    }

    @Override
    public Map<String, Object> getMapFromTemplateModel(String expression, TemplateModel model) {
        Map<String, Object> mapResult = new HashMap<>();
        Map map = (LazyMap) model.getValue(expression);
        List<String> keyList = (List<String>) map.keySet().stream().collect(Collectors.toList());

        for (String key: keyList) {
            Object value = map.get(key);
            if(value != null && value instanceof String){
                mapResult.put(key, value);
            }

        }

        return mapResult;
    }
}
