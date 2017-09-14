package uk.gov.education.honours;

import org.junit.Test;

import static org.junit.Assert.*;

public class KissflowApiTest {
    @Test
    public void getShortlist() throws Exception {
        KissflowApi kissflowApi = new KissflowApi("7d426333-915b-11e7-9ddd-b1297a8fe2e3");
        kissflowApi.getShortlist("","2018 NY");
    }

}