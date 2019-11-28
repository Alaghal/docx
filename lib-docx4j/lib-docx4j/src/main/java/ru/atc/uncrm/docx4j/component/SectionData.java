package ru.atc.uncrm.docx4j.component;



import lombok.Getter;

import java.util.List;
import java.util.Map;

@Getter
public class SectionData {
    private List<Map<String, Object>> dataToAddForTable;
    private Map<String, Object> dataToAddForParagraphs;

    public SectionData( List<Map<String, Object>> dataToAddForTable, Map<String, Object> dataToAddForParagraps) {
        this.dataToAddForTable = dataToAddForTable;
        this.dataToAddForParagraphs = dataToAddForParagraps;
    }
}
