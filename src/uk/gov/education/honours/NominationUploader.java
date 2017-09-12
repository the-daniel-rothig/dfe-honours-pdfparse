package uk.gov.education.honours;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.fit.pdfdom.PDFDomTree;
import org.json.simple.parser.ParseException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;
import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class NominationUploader {

    private final KissflowApi kissflow;
    private final FileUploader fileUploader;

    public NominationUploader(String kissflowApiKey, String fileApiKey) {
        kissflow = new KissflowApi(kissflowApiKey);
        fileUploader = new FileUploader(fileApiKey);
    }

    private static final List<String> boilerplate = Arrays.asList(
            "You have a new nomination submission",
            "View full details",
            "-- Cabinet Office Honours and Appointments Secretariat 1 Horse Guards Road London null SW1A 2HQ United Kingdom null null T: 020 7276 2777");


    public XSSFWorkbook getShortlist(String directorate, String round) throws IOException, ParseException {
        return kissflow.getShortlist(directorate, round);
    }

    public void importShortlist(File shortlistFile) throws IOException, ParseException, InvalidFormatException {
        kissflow.importShortlist(shortlistFile);
    }

    public XSSFWorkbook getFinalShortlist(String round) throws IOException, ParseException {
        return kissflow.getFinalShortlist(round);
    }

    public Result uploadNomination(Path nominationPdf, String targetFileName, Map<String,File> fileBucket) throws Exception {
        String filename = nominationPdf.toAbsolutePath().toString();
        List<Element> words = getWordsFromPdf(filename);
        List<Phrase> allPhrases = bunchWordsIntoPhrases(words);

        List<Section> res = structurePhrasesIntoSections(allPhrases);

        List<FileNameAndPath> evidenceFiles = filterToEvidenceFiles(res, fileBucket);
        List<FileNameAndPath> uploadedFiles = new ArrayList<>();
        for (FileNameAndPath fnap : evidenceFiles) {
            String url = fileUploader.sendFile(fnap.path, fnap.name);
            uploadedFiles.add(new FileNameAndPath(fnap.name, url));
        }

        String nom = fileUploader.sendFile(filename, targetFileName);
        String nextStepUri = kissflow.sendToKissflow(res, targetFileName, nom, uploadedFiles);

        Matcher matcher = Pattern.compile("Honours nomination web form submitted for (.+)\\.pdf").matcher(targetFileName);


        return new Result(evidenceFiles.size(), matcher.find() ? matcher.group(1) : targetFileName, nextStepUri);
    }

    public static class Result {
        public final boolean success = true; // no error handling, so if the function returns it's successful
        public final int supportDocumentCount;
        public final String nominationFileName;
        public final String nextStepUri;

        public Result(int supportDocumentCount, String nominationFileName, String nextStepUri) {
            this.supportDocumentCount = supportDocumentCount;
            this.nominationFileName = nominationFileName;
            this.nextStepUri = nextStepUri;
        }
    }

    private static List<FileNameAndPath> filterToEvidenceFiles(List<Section> res, Map<String, File> fileBucket) {
        List<FileNameAndPath> fnaps = new ArrayList<>();
        for(Map<String, String> x : res.get(2).questions.get("Evidence of your nominee's contribution").answers) {
            File file = fileBucket.getOrDefault(x.get("Attachment name"), null);
            if (file != null) fnaps.add(new FileNameAndPath(x.get("AttachmentName"), file.getAbsolutePath()));
        }
        for(Map<String, String> x : res.get(2).questions.get("Letters of support").answers) {
            File file = fileBucket.getOrDefault(x.get("Attachment name"), null);
            if (file != null) fnaps.add(new FileNameAndPath(x.get("AttachmentName"), file.getAbsolutePath()));
        }
        return fnaps;
    }

    static class FileNameAndPath {
        String name;
        String path;

        FileNameAndPath(String name, String absolutePath) {
            this.name = name;
            this.path = absolutePath;
        }
    }


    private static List<Element> getWordsFromPdf(String filename) throws IOException, ParserConfigurationException, XPathExpressionException {
        PDDocument pdf = PDDocument.load(new java.io.File(filename));
        PDFDomTree parser = new PDFDomTree();
        Document xmlDoc = parser.createDOM(pdf);

        XPath xPath = XPathFactory.newInstance().newXPath();
        NodeList nodes = (NodeList)xPath.evaluate("//div[contains(concat(\" \", normalize-space(@class), \" \"), \" p \")]", xmlDoc.getDocumentElement(), XPathConstants.NODESET);


        List<Element> res = new ArrayList<>();
        for (int i = 0; i < nodes.getLength(); i++) {
            res.add((Element) nodes.item(i));
        }
        return res;
    }

    private static List<Phrase> bunchWordsIntoPhrases(List<Element> words) {
        List<Phrase> allPhrases = new ArrayList<>();
        List<String> currentPhrase = new ArrayList<>();
        WordParams lastWord = null;
        WordParams firstWordOfCurrentPhrase = null;


        for (Element e : words) {
            WordParams wp = WordParams.parse(e.getAttribute("style"));

            if (!currentPhrase.isEmpty() && !wp.continues(lastWord) && lastWord != null && firstWordOfCurrentPhrase != null) {
                allPhrases.add(new Phrase(firstWordOfCurrentPhrase, lastWord, String.join(" ",currentPhrase)));
                currentPhrase.clear();
            }

            if (currentPhrase.isEmpty()) firstWordOfCurrentPhrase = wp;

            currentPhrase.add(e.getTextContent());
            lastWord =wp;
        }

        allPhrases.add(new Phrase(firstWordOfCurrentPhrase, lastWord, String.join(" ",currentPhrase)));
        return allPhrases;
    }


    private static List<Section> structurePhrasesIntoSections(List<Phrase> allPhrases) {
        List<Section> res = new ArrayList<>();
        Section currentSection = null;

        List<Phrase> currentLabels = new ArrayList<>();
        boolean currentlyGatheringHeaders = false;
        Phrase previousValuePhrase = null;


        for (Phrase phrase : allPhrases) {
            if (boilerplate.contains(phrase.phrase)) continue;

            if (phrase.keyWord.getType() == 0) {
                OptionalInt bestFittingLabel = IntStream.range(0, currentLabels.size())
                        .filter(i -> phrase.keyWord.horizontallyAlinged(currentLabels.get(i).keyWord))
                        .reduce((a, b) -> b);

                if (!bestFittingLabel.isPresent()) {
                    currentSection.getCurrentQuestion().addSimpleAnswer(phrase.phrase);
                }
                else {
                    if (previousValuePhrase != null && !phrase.onSameLineAs(previousValuePhrase)) {
                        currentSection.getCurrentQuestion().addAnswer();
                    }
                    currentSection.getCurrentQuestion().addToLastAnswer(currentLabels.get(bestFittingLabel.getAsInt()).phrase, phrase.phrase);
                    previousValuePhrase = phrase;
                }
            } else {
                previousValuePhrase = null;
            }

            if (phrase.keyWord.getType() == 1) {
                if (!currentlyGatheringHeaders) {
                    currentLabels.clear();
                }
                currentLabels.add(phrase);
                currentlyGatheringHeaders = true;
            } else {
                currentlyGatheringHeaders = false;
            }

            if (phrase.keyWord.getType() == 2) {
                currentSection.addQuestion(phrase.phrase);
                currentLabels.clear();
            }

            if (phrase.keyWord.getType() == 3) {
                currentSection = new Section(phrase.phrase);
                res.add(currentSection);
                currentLabels.clear();

            }
        }
        return res;
    }

}

class Section {
    String label;
    Map<String, Question> questions = new HashMap<>();

    private Question currentQuestion = null;

    Section(String label) {
        this.label = label;
    }

    Question addQuestion(String label) {
        Question q = new Question();
        q.label = label;
        questions.put(label, q);
        currentQuestion = q;
        return q;
    }

    Question getCurrentQuestion() {
        if (currentQuestion == null) addQuestion("Details");
        return currentQuestion;
    }

    public Map<String,String> props() {
        return props("Details");
    }

    public Map<String,String> props(String name) {
        return questions.get(name).answers.get(0);
    }

    static class Question {
        String label;
        List<Map<String,String>> answers = new ArrayList<>();
        String simpleAnswer = null;
        private  Map<String, String> lastAnswer = null;

        void addAnswer() throws UnsupportedOperationException {
            if (simpleAnswer != null) throw new UnsupportedOperationException("Parse error: question with simple answer receiving complex answer");

            Map<String, String> a = new HashMap<>();
            answers.add(a);
            lastAnswer = a;
        }

        void addToLastAnswer(String key, String value) {
            if (lastAnswer == null) addAnswer();
            lastAnswer.put(key, value);
        }

        void addSimpleAnswer(String answer) throws UnsupportedOperationException {
            if (!answers.isEmpty()) throw new UnsupportedOperationException("Parse error: question with complex answer receiving simple answer");

            simpleAnswer = simpleAnswer == null ? answer : simpleAnswer + "\n" + answer;
        }

        public String toString() {
            return String.format("= QUESTION: %s =\n%s\n", label,
                    answers.isEmpty() ? simpleAnswer : answers.stream().map(x -> x.entrySet().stream().map(y -> String.format("%s: %s", y.getKey(), y.getValue())).collect(Collectors.joining(", "))).collect(Collectors.joining("\n")));
        }
    }

    public String toString() {
        return String.format("== SECTION: %s ==\n%s\n\n", label,
                questions.values().stream().map(Object::toString).collect(Collectors.joining("\n")));
    }
}

class Phrase {
    WordParams keyWord;
    private WordParams lastWord;
    String phrase;

    Phrase(WordParams keyWord, WordParams lastWord, String phrase) {
        this.keyWord = keyWord;
        this.lastWord = lastWord;
        this.phrase = phrase;
    }

    boolean onSameLineAs(Phrase previousValuePhrase) {
        if (previousValuePhrase == null) return false;

        // assumes centered vertical alignment

        return Math.abs(keyWord.top + lastWord.top - (previousValuePhrase.keyWord.top + previousValuePhrase.lastWord.top)) < 0.1;
    }

    @Override
    public String toString() {
        return phrase;
    }
}

class WordParams {
    Double fontSize;
    String fontWeight;
    Double top;
    Double left;
    Double width;

    static WordParams parse(String style) {
        WordParams res = new WordParams();
        res.fontSize = Double.parseDouble(findOr(fontSizePattern, style, "0"));
        res.fontWeight = findOr(fontWeightPattern, style, "normal");
        res.top = Double.parseDouble(findOr(topPattern, style, "0"));
        res.left = Double.parseDouble(findOr(leftPattern, style, "0"));
        res.width = Double.parseDouble(findOr(widthPattern, style, "0"));

        return res;
    }

    boolean continues(WordParams previousWord) {
        if (previousWord == null) return false;

        if (!vagueEq(fontSize, previousWord.fontSize) || !Objects.equals(fontWeight, previousWord.fontWeight)) return false;

        //continuation on the same line
        if (vagueEq(top, previousWord.top) && left - previousWord.left - previousWord.width < 10) return true;

        //line break
        if (top >  1 + previousWord.top && top < 20 + previousWord.top) return true;

        //false otherwise; this includes page breaks.
        return false;
    }

    boolean horizontallyAlinged(WordParams aboveWord) {
        return vagueEq(left, aboveWord.left);
    }

    // 0 = data
    // 1 = label
    // 2 = question
    // 3 = section
    int getType() {
        if (fontSize > 15) return 3;
        else if (fontSize > 13) return 2;
        else if (Objects.equals(fontWeight, "normal")) return 0;
        else return 1;
    }

    private static String findOr(Pattern p, String text, String valIfNotFound) {
        Matcher matcher = p.matcher(text);
        if (matcher.find()) {
            return matcher.group(1);
        } else {
            return valIfNotFound;
        }
    }

    private boolean vagueEq(Double x, Double y) {
        return Math.abs(x - y) < 0.001;
    }

    private static Pattern fontSizePattern = Pattern.compile("font-size:([\\.0-9]+)pt");
    private static Pattern fontWeightPattern = Pattern.compile("font-weight:([^;]+)");
    private static Pattern topPattern = Pattern.compile("top:([\\.0-9]+)pt");
    private static Pattern leftPattern = Pattern.compile("left:([\\.0-9]+)pt");
    private static Pattern widthPattern = Pattern.compile("width:([\\.0-9]+)pt");
}


