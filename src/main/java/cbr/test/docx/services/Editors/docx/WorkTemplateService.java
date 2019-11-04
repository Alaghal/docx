package cbr.test.docx.services.Editors.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public interface WorkTemplateService {

    public WordprocessingMLPackage getTemplate(String name) ;
}
