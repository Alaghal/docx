package cbr.test.docx.services.Editors.docx;

import lombok.Getter;

import java.util.List;
import java.util.Map;

@Getter
public class SectionData {
    private List<Map<String, String>> dataToAddForTable;
    private Map<String, String> dataToAddForParagraps;

    public SectionData( List<Map<String, String>> dataToAddForTable, Map<String, String> dataToAddForParagraps) {
        this.dataToAddForTable = dataToAddForTable;
        this.dataToAddForParagraps = dataToAddForParagraps;
    }
}
