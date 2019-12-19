package ch.bbcag.fit4ipa;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.NameValuePair;
import org.apache.http.client.HttpClient;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.message.BasicNameValuePair;
import org.apache.poi.xwpf.usermodel.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.w3c.dom.Text;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

public class Appendix {
    private Map<List<String>, String> filetypes;

    private boolean isDebug = false;
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

    public static void main(String[] args) throws Exception {
        if (new File(filename).exists()) {
            Files.delete(new File(filename).toPath());
        }
        var a = new Appendix();
        a.writeWordDocument();
        Desktop.getDesktop().open(new File(filename));
    }


    private Appendix() {
        initFiletypes();
        initSwing();
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
        try {
            var rawCode = Files.readString(file.toPath());
            var lexer = getLexerForFilename(file.getName());

            System.out.println(lexer + "\t" + file.getName());

            var htmlCode = getHtmlFromCode(rawCode, lexer);

            var titleParagraph = document.createParagraph();
            titleParagraph.setPageBreak(true);

            var title = titleParagraph.createRun();
            title.setText(file.getAbsolutePath().replace(root + "\\", ""));
            title.setBold(true);
            title.setFontSize(14);

            var codeParagraph = document.createParagraph();
            writeCodeToParagraph(codeParagraph, htmlCode);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void writeCodeToParagraph(XWPFParagraph paragraph, String code) {
        code = code.replace("\n", "<br/>");
        code = code.replaceAll("</span>([^<br>].*?)<span", "</span><span>$1</span><span");
        var doc = Jsoup.parse(code);
        var root = doc.select("pre").get(0);

        root.childNodes().forEach(node -> handleNode(paragraph, node));
    }

    private void handleNode(XWPFParagraph paragraph, Node node) {
        if (node.childNodes().size() > 0) {
            node.childNodes().forEach(n -> handleNode(paragraph, n));
        } else {
            if (node instanceof TextNode) {
                writeTextNode(paragraph.createRun(), (TextNode) node);
            } else if (node instanceof Element) {
                Element e = (Element) node;
                if (e.tagName().equals("br")) {
                    paragraph.createRun().addBreak(BreakType.TEXT_WRAPPING);
                }
            }
        }
    }

    private void writeTextNode(XWPFRun run, TextNode node) {
        run.setFontFamily("Lucida Console");
        run.setFontSize(8);

        run.setText(node.getWholeText());
        run.setColor(getColorFromElement(node.parentNode()));
        run.setBold(isBold(node));
        run.setItalic(isItalic(node));
        run.setStrikeThrough(isStrikethrough(node));

        if (isUnderline(node)) {
            run.setUnderline(UnderlinePatterns.DASH);
            run.setUnderlineColor(run.getColor());
        }
        var backgroundColour = getBackgroundColor(node);
        if (backgroundColour != null) {
            CTShd cTShd = run.getCTR().addNewRPr().addNewShd();
            cTShd.setVal(STShd.CLEAR);
            cTShd.setColor("auto");
            cTShd.setFill(backgroundColour);
        }
    }

    private String getColorFromElement(Node node) {
        String[] styles = node.attr("style").split(";");
        for (String s : styles) {
            if (s.startsWith("color:")) {
                return s.split(":")[1].replace("#", "");
            }
        }
        return "000000";
    }

    private boolean isBold(Node node) {
        String[] styles = node.attr("style").split(";");
        for (String s : styles) {
            if (s.startsWith("font-weight:")) {
                String value = s.split(":")[1];
                return value.equals("bold");
            }
        }
        return false;
    }

    private boolean isItalic(Node node) {
        String[] styles = node.attr("style").split(";");
        for (String s : styles) {
            if (s.startsWith("font-style:")) {
                String value = s.split(":")[1];
                return value.equals("italic") || value.equals("oblique");
            }
        }
        return false;
    }

    private boolean isUnderline(Node node) {
        String[] styles = node.attr("style").split(";");
        for (String s : styles) {
            if (s.startsWith("text-decoration:")) {
                String value = s.split(":")[1];
                return value.contains("underline");
            }
        }
        return false;
    }

    private boolean isStrikethrough(Node node) {
        String[] styles = node.attr("style").split(";");
        for (String s : styles) {
            if (s.startsWith("text-decoration:")) {
                String value = s.split(":")[1];
                return value.contains("line-through");
            }
        }
        return false;
    }

    private String getBackgroundColor(Node node) {
        String[] styles = node.attr("style").split(";");
        for (String s : styles) {
            if (s.startsWith("background-color:")) {
                return s.split(":")[1].replace("#", "");
            }
        }
        return null;
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
            return "D:\\code\\fit4ipa\\Appendix1";
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

    private String getHtmlFromCode(String raw, String lexer) throws IOException {
        HttpClient httpclient = HttpClients.createDefault();
        HttpPost httppost = new HttpPost("http://hilite.me/api");

        List<NameValuePair> params = new ArrayList<>(2);
        params.add(new BasicNameValuePair("code", raw.trim()));
        params.add(new BasicNameValuePair("lexer", lexer));
        params.add(new BasicNameValuePair("style", "vs"));
        httppost.setEntity(new UrlEncodedFormEntity(params, "UTF-8"));

        HttpResponse response = httpclient.execute(httppost);
        HttpEntity entity = response.getEntity();

        if (entity != null && response.getStatusLine().getStatusCode() == 200) {
            return IOUtils.toString(entity.getContent(), StandardCharsets.UTF_8).replace("<!-- HTML generated using hilite.me -->", "");
        } else {
            if (entity != null && entity.getContent() != null) {
                System.err.println(IOUtils.toString(entity.getContent(), StandardCharsets.UTF_8));
            }
            throw new RuntimeException("Failed to generate html");
        }
    }

    private String getLexerForFilename(String filename) {
        if (filename.contains(".")) {
            String[] parts = filename.split("\\.");
            filename = "." + parts[parts.length - 1];
        }
        for (List<String> suffixes : this.filetypes.keySet()) {
            if (suffixes.contains(filename)) {
                return this.filetypes.get(suffixes);
            }
        }
        return "text";
    }

    @SuppressWarnings("SpellCheckingInspection")
    private void initFiletypes() {
        this.filetypes = new HashMap<>();
        this.filetypes.put(List.of(".json"), "json");
        this.filetypes.put(List.of(".js"), "js");
        this.filetypes.put(List.of(".ts"), "ts");
        this.filetypes.put(List.of(".css"), "css");
        this.filetypes.put(List.of(".html"), "html");
        this.filetypes.put(List.of(".md"), "text");
        this.filetypes.put(List.of(".croc".split(",")), "croc".split(",")[0]);
        this.filetypes.put(List.of(".dg".split(",")), "dg".split(",")[0]);
        this.filetypes.put(List.of(".factor".split(",")), "factor".split(",")[0]);
        this.filetypes.put(List.of(".fy,.fancypack".split(",")), "fancy,fy".split(",")[0]);
        this.filetypes.put(List.of(".io".split(",")), "io".split(",")[0]);
        this.filetypes.put(List.of(".lua,.wlua".split(",")), "lua".split(",")[0]);
        this.filetypes.put(List.of(".moon".split(",")), "moon,moonscript".split(",")[0]);
        this.filetypes.put(List.of(".pl,.pm".split(",")), "perl,pl".split(",")[0]);
        this.filetypes.put(List.of(".py3tb".split(",")), "py3tb".split(",")[0]);
        this.filetypes.put(List.of(".py,.pyw,.sc,SConstruct,SConscript,.tac,.sage".split(",")), "python,py,sage".split(",")[0]);
        this.filetypes.put(List.of(".pytb".split(",")), "pytb".split(",")[0]);
        this.filetypes.put(List.of(".rb,.rbw,Rakefile,.rake,.gemspec,.rbx,.duby".split(",")), "rb,ruby,duby".split(",")[0]);
        this.filetypes.put(List.of(".tcl".split(",")), "tcl".split(",")[0]);
        this.filetypes.put(List.of(".c-objdump".split(",")), "c-objdump".split(",")[0]);
        this.filetypes.put(List.of(".s".split(",")), "ca65".split(",")[0]);
        this.filetypes.put(List.of(".cpp-objdump,.c++-objdump,.cxx-objdump".split(",")), "cpp-objdump,c++ - objdumb,cxx - objdump".split(",")[0]);
        this.filetypes.put(List.of(".d-objdump".split(",")), "d-objdump".split(",")[0]);
        this.filetypes.put(List.of(".s,.S".split(",")), "gas".split(",")[0]);
        this.filetypes.put(List.of(".ll".split(",")), "llvm".split(",")[0]);
        this.filetypes.put(List.of(".asm,.ASM".split(",")), "nasm".split(",")[0]);
        this.filetypes.put(List.of(".objdump".split(",")), "objdump".split(",")[0]);
        this.filetypes.put(List.of(".adb,.ads,.ada".split(",")), "ada,ada95ada2005".split(",")[0]);
        this.filetypes.put(List.of(".bmx".split(",")), "blitzmax,bmax".split(",")[0]);
        this.filetypes.put(List.of(".c,.h,.idc".split(",")), "c".split(",")[0]);
        this.filetypes.put(List.of(".cbl,.CBL".split(",")), "cobolfree".split(",")[0]);
        this.filetypes.put(List.of(".cob,.COB,.cpy,.CPY".split(",")), "cobol".split(",")[0]);
        this.filetypes.put(List.of(".cpp,.hpp,.c++,.h++,.cc,.hh,.cxx,.hxx,.C,.H,.cp,.CPP".split(",")), "cpp,c++".split(",")[0]);
        this.filetypes.put(List.of(".cu,.cuh".split(",")), "cuda,cu".split(",")[0]);
        this.filetypes.put(List.of(".pyx,.pxd,.pxi".split(",")), "cython,pyx".split(",")[0]);
        this.filetypes.put(List.of(".d,.di".split(",")), "d".split(",")[0]);
        this.filetypes.put(List.of(".pas".split(",")), "delphi,pas,pascal,objectpascal".split(",")[0]);
        this.filetypes.put(List.of(".dylan,.dyl,.intr".split(",")), "dylan".split(",")[0]);
        this.filetypes.put(List.of(".lid,.hdp".split(",")), "dylan-lid,lid".split(",")[0]);
        this.filetypes.put(List.of(".ec,.eh".split(",")), "ec".split(",")[0]);
        this.filetypes.put(List.of(".fan".split(",")), "fan".split(",")[0]);
        this.filetypes.put(List.of(".flx,.flxh".split(",")), "felix,flx".split(",")[0]);
        this.filetypes.put(List.of(".f,.f90,.F,.F90".split(",")), "fortran".split(",")[0]);
        this.filetypes.put(List.of(".vert,.frag,.geo".split(",")), "glsl".split(",")[0]);
        this.filetypes.put(List.of(".go".split(",")), "go".split(",")[0]);
        this.filetypes.put(List.of(".def,.mod".split(",")), "modula2,m2".split(",")[0]);
        this.filetypes.put(List.of(".monkey".split(",")), "monkey".split(",")[0]);
        this.filetypes.put(List.of(".nim,.nimrod".split(",")), "nimrod,nim".split(",")[0]);
        this.filetypes.put(List.of(".m,.h".split(",")), "objective-c,objectivec,obj - c,objc".split(",")[0]);
        this.filetypes.put(List.of(".mm,.hh".split(",")), "objective-c++,objectivec++,obj - c++,objc++".split(",")[0]);
        this.filetypes.put(List.of(".ooc".split(",")), "ooc".split(",")[0]);
        this.filetypes.put(List.of(".prolog,.pro,.pl".split(",")), "prolog".split(",")[0]);
        this.filetypes.put(List.of(".rs,.rc".split(",")), "rust".split(",")[0]);
        this.filetypes.put(List.of(".vala,.vapi".split(",")), "vala,vapi".split(",")[0]);
        this.filetypes.put(List.of(".smali".split(",")), "smali".split(",")[0]);
        this.filetypes.put(List.of(".boo".split(",")), "boo".split(",")[0]);
        this.filetypes.put(List.of(".aspx,.asax,.ascx,.ashx,.asmx,.axd".split(",")), "aspx-cs".split(",")[0]);
        this.filetypes.put(List.of(".cs".split(",")), "csharp,c#".split(",")[0]);
        this.filetypes.put(List.of(".fs,.fsi".split(",")), "fsharp".split(",")[0]);
        this.filetypes.put(List.of(".n".split(",")), "nemerle".split(",")[0]);
        this.filetypes.put(List.of(".aspx,.asax,.ascx,.ashx,.asmx,.axd".split(",")), "aspx-vb".split(",")[0]);
        this.filetypes.put(List.of(".vb,.bas".split(",")), "vb.net,vbnet".split(",")[0]);
        this.filetypes.put(List.of(".PRG,.prg".split(",")), "Clipper,XBase".split(",")[0]);
        this.filetypes.put(List.of(".cl,.lisp,.el".split(",")), "common-lisp,cl".split(",")[0]);
        this.filetypes.put(List.of(".v".split(",")), "coq".split(",")[0]);
        this.filetypes.put(List.of(".ex,.exs".split(",")), "elixir,ex,exs".split(",")[0]);
        this.filetypes.put(List.of(".erl,.hrl,.es,.escript".split(",")), "erlang".split(",")[0]);
        this.filetypes.put(List.of(".erl-sh".split(",")), "erl".split(",")[0]);
        this.filetypes.put(List.of(".hs".split(",")), "haskell,hs".split(",")[0]);
        this.filetypes.put(List.of(".kk,.kki".split(",")), "koka".split(",")[0]);
        this.filetypes.put(List.of(".lhs".split(",")), "lhs,literate-haskell".split(",")[0]);
        this.filetypes.put(List.of(".lsp,.nl".split(",")), "newlisp".split(",")[0]);
        this.filetypes.put(List.of(".ml,.mli,.mll,.mly".split(",")), "ocaml".split(",")[0]);
        this.filetypes.put(List.of(".opa".split(",")), "opa".split(",")[0]);
        this.filetypes.put(List.of(".rkt,.rktl".split(",")), "racket,rkt".split(",")[0]);
        this.filetypes.put(List.of(".sml,.sig,.fun".split(",")), "sml".split(",")[0]);
        this.filetypes.put(List.of(".scm,.ss".split(",")), "scheme,scm".split(",")[0]);
        this.filetypes.put(List.of(".sv,.svh".split(",")), "sv".split(",")[0]);
        this.filetypes.put(List.of(".v".split(",")), "v".split(",")[0]);
        this.filetypes.put(List.of(".vhdl,.vhd".split(",")), "vhdl".split(",")[0]);
        this.filetypes.put(List.of(".aj".split(",")), "aspectj".split(",")[0]);
        this.filetypes.put(List.of(".ceylon".split(",")), "ceylon".split(",")[0]);
        this.filetypes.put(List.of(".clj".split(",")), "clojure,clj".split(",")[0]);
        this.filetypes.put(List.of(".gs,.gsx,.gsp,.vark".split(",")), "gosu".split(",")[0]);
        this.filetypes.put(List.of(".gst".split(",")), "gst".split(",")[0]);
        this.filetypes.put(List.of(".groovy".split(",")), "groovy".split(",")[0]);
        this.filetypes.put(List.of(".ik".split(",")), "ioke,ik".split(",")[0]);
        this.filetypes.put(List.of(".java".split(",")), "java".split(",")[0]);
        this.filetypes.put(List.of(".kt".split(",")), "kotlin".split(",")[0]);
        this.filetypes.put(List.of(".scala".split(",")), "scala".split(",")[0]);
        this.filetypes.put(List.of(".xtend".split(",")), "xtend".split(",")[0]);
        this.filetypes.put(List.of(".bug".split(",")), "bugs,winbugs,openbugs".split(",")[0]);
        this.filetypes.put(List.of(".pro".split(",")), "idl".split(",")[0]);
        this.filetypes.put(List.of(".jag,.bug".split(",")), "jags".split(",")[0]);
        this.filetypes.put(List.of(".jl".split(",")), "julia,jl".split(",")[0]);
        this.filetypes.put(List.of(".m".split(",")), "matlab".split(",")[0]);
        this.filetypes.put(List.of(".mu".split(",")), "mupad".split(",")[0]);
        this.filetypes.put(List.of(".m".split(",")), "octave".split(",")[0]);
        this.filetypes.put(List.of(".Rout".split(",")), "rconsole,rout".split(",")[0]);
        this.filetypes.put(List.of(".Rd".split(",")), "rd".split(",")[0]);
        this.filetypes.put(List.of(".S,.R,.Rhistory,.Rprofile".split(",")), "splus,s,r".split(",")[0]);
        this.filetypes.put(List.of(".sci,.sce,.tst".split(",")), "scilab".split(",")[0]);
        this.filetypes.put(List.of(".stan".split(",")), "stan".split(",")[0]);
        this.filetypes.put(List.of(".abap".split(",")), "abap".split(",")[0]);
        this.filetypes.put(List.of(".applescript".split(",")), "applescript".split(",")[0]);
        this.filetypes.put(List.of(".asy".split(",")), "asy,asymptote".split(",")[0]);
        this.filetypes.put(List.of(".au3".split(",")), "autoit,Autoit".split(",")[0]);
        this.filetypes.put(List.of(".ahk,.ahkl".split(",")), "ahk".split(",")[0]);
        this.filetypes.put(List.of(".awk".split(",")), "awk,gawk,mawk".split(",")[0]);
    }

}
