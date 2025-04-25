package tpm.excel;


import lombok.Getter;
import lombok.Setter;


public class ItemField {
    @Setter
    @Getter
    private long barcode;
    @Setter
    @Getter
    private String productName;
    @Setter
    @Getter
    private String description;
    @Setter
    @Getter
    private String nutrition;
    @Setter
    @Getter
    private String brand;
    @Setter
    @Getter
    private int expiration;
    @Setter
    @Getter
    private String storageCondition;
    @Setter
    @Getter
    private String composition;
    @Setter
    @Getter
    private String image;
}
