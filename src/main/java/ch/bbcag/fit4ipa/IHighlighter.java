package ch.bbcag.fit4ipa;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;

public interface IHighlighter {

    void highlightToParagraph(File file, XWPFParagraph paragraph);

}
