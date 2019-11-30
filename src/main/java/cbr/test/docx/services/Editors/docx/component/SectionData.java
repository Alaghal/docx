package cbr.test.docx.services.Editors.docx.component;



import lombok.Getter;

import java.util.Map;

@Getter
public class SectionData {
    private Map<String, Object> data;

    public SectionData( Map<String, Object> data) {
        this.data = data;
    }
}
