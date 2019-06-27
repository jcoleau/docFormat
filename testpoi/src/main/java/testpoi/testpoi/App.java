package testpoi.testpoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

public class App {

	public static void main(String[] args) {
		try {
			XWPFDocument doc = new XWPFDocument();
			CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
			XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(doc, sectPr);
			FileOutputStream out = new FileOutputStream(new File("C:/Users/yogac/documents/testFile.docx"));

			CTP ctpHeader = CTP.Factory.newInstance();
			CTR ctrHeader = ctpHeader.addNewR();
			CTText ctHeader = ctrHeader.addNewT();
			String headerText = "The best Company in the world";

			ctHeader.setStringValue(headerText);
			XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, doc);
			XWPFParagraph[] parsHeader = new XWPFParagraph[1];
			parsHeader[0] = headerParagraph;
			policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);

			XWPFParagraph bodyParagraph = doc.createParagraph();
			bodyParagraph.setAlignment(ParagraphAlignment.CENTER);
			XWPFRun r = bodyParagraph.createRun();
			r.setBold(true);
			r.setText("I was centered in eclipse. ");
			r.addBreak();

			r.setText("Another text?");

			doc.write(out);
			out.close();
		
		} catch (Exception e) {
			e.getStackTrace();
		}
		System.out.println("Done");
	}

}
