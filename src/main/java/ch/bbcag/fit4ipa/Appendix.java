package ch.bbcag.fit4ipa;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Collection;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

public class Appendix {

    private static String filename = "Anhang 1.docx";
    private static List<String> filesToIgnore = List.of(
            ".idea",
            "e2e",
            ".ico",
            "test.ts",
            ".png",
            ".gif",
            ".iml",
            "package-lock.json",
            ".bmpr",
            ".sln",
            ".user",
            ".vs",
            "build",
            ".mailmap",
            ".vsdx",
            ".docx",
            ".docx",
            ".xls",
            ".xlsx",
            ".git",
            ".vsproj",
            "node_modules",
            "gradle",
            "obj",
            ".dll",
            ".cache",
            ".csproj",
            "bin");

    @SuppressWarnings("FieldCanBeLocal")
    private final boolean isDebug = true;
    @SuppressWarnings("FieldCanBeLocal")
    private final boolean useLocalHighlighter = false;
    private final IHighlighter highlighter;

    private Appendix() {
        initSwing();
        if (useLocalHighlighter) {
            highlighter = null;
        } else {
            highlighter = new HtmlHighlighter();
        }
    }

    public static void main(String[] args) throws Exception {
        if (new File(filename).exists()) {
            Files.delete(new File(filename).toPath());
        }
        var a = new Appendix();
        a.writeWordDocument();
        Desktop.getDesktop().open(new File(filename));
    }

    private void writeWordDocument() {
        var document = new XWPFDocument();

        String root = getPath();
        getFileToPaste(root).forEach(file -> writeCodeToDoc(document, file, root));

        try (FileOutputStream out = new FileOutputStream(filename)) {
            document.write(out);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void writeCodeToDoc(XWPFDocument document, File file, String root) {
        var titleParagraph = document.createParagraph();
        titleParagraph.setPageBreak(true);

        var title = titleParagraph.createRun();
        title.setText(file.getAbsolutePath().replace(root + "\\", ""));
        title.setBold(true);
        title.setFontSize(14);

        var codeParagraph = document.createParagraph();
        highlighter.highlightToParagraph(file, codeParagraph);
    }


    private void initSwing() {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | UnsupportedLookAndFeelException e) {
            e.printStackTrace();
        }
    }

    private String getPath() {
        if (isDebug) {
            return "/home/benji/src/BenjaminRaison/gen-appendix/src";
        }
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooser.setDialogTitle("Select IPA root folder");
        chooser.setDialogType(JFileChooser.OPEN_DIALOG);
        chooser.setAcceptAllFileFilterUsed(false);

        if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
            return chooser.getSelectedFile().getAbsolutePath();
        } else {
            System.exit(1);
        }
        return null;
    }

    private List<File> getFileToPaste(String path) {
        Collection<File> files = FileUtils.listFiles(new File(path), null, true);
        return files.parallelStream().filter(this::shouldIncludeFile).sorted(Comparator.comparing(File::getAbsolutePath)).collect(Collectors.toList());
    }

    private boolean shouldIncludeFile(File file) {
        if (file.isDirectory()) {
            return false;
        }
        for (String part : filesToIgnore) {
            if (file.getAbsolutePath().contains(part)) {
                return false;
            }
        }
        return true;
    }
}
