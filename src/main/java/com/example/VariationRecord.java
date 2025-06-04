package com.example;

public class VariationRecord {
    private final String variationName;
    private final String option1;
    private final String option2;
    private final String productCode;

    public VariationRecord(String variationName, String option1, String option2, String productCode) {
        this.variationName = variationName;
        this.option1 = option1;
        this.option2 = option2;
        this.productCode = productCode;
    }

    public String getVariationName() { return variationName; }
    public String getOption1() { return option1; }
    public String getOption2() { return option2; }
    public String getProductCode() { return productCode; }
}
