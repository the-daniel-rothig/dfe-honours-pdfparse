package uk.gov.education.honours;

import org.junit.Test;

import java.io.File;
import java.io.InputStream;
import java.io.RandomAccessFile;
import java.nio.channels.Channels;
import java.nio.file.Files;

import static org.junit.Assert.*;

public class KissflowApiTest {
    @Test
    public void getShortlist() throws Exception {
        KissflowApi kissflowApi = new KissflowApi("7d426333-915b-11e7-9ddd-b1297a8fe2e3");
        kissflowApi.getShortlist("","2018 NY");
    }

    @Test
    public void largeFile() throws Exception {
        FileUploader fileUploader = new FileUploader("e192af2487358543335a");
        RandomAccessFile f = new RandomAccessFile("t", "rw");
        f.setLength(99 * 1024 * 1024);

        InputStream inputStream = Channels.newInputStream(f.getChannel());
        fileUploader.sendFile(inputStream, Files.probeContentType(new File("massive.txt").toPath()), "massive.txt");
    }

}