//package cbr.test.docx.services.Editors.docx.factory.impl;
//
//import cbr.test.docx.services.Editors.docx.factory.FactoryTemplateWriterDataSection;
//import lombok.Getter;
////import ru.atc.template.model.TemplateModel;
////import ru.atc.uncrm.docx4j.component.WriterDataSectionsForTemplates;
////import ru.atc.uncrm.docx4j.component.impl.DkiduWordWriterDataSectionsForTemplates;
//
//import java.io.OutputStream;
//
//@Getter
//public class DkiduFactoryTemplateWriterDataSection implements FactoryTemplateWriterDataSection {
//
//    private OutputStream outputStream;
//    private String codeReport;
// //   private TemplateModel model;
//
//
//
//    public DkiduFactoryTemplateWriterDataSection(String codeReport, TemplateModel model, OutputStream outputStream) {
//        this.outputStream = outputStream;
//        this.codeReport= codeReport;
//       // this.model = model;
//    }
//
//    @Override
//    public WriterDataSectionsForTemplates createTemplateWriteDataSection() {
//
//        int u = 0;
//       if(codeReport.equals("AZ_UK_DKIDU")){
//             return  new DkiduWordWriterDataSectionsForTemplates(codeReport, model);
//       }
//
//       return null;
//    }
//}
