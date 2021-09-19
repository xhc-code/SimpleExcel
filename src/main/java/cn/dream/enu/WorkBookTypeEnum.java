package cn.dream.enu;

public enum WorkBookTypeEnum {

    XLS("xls"),XLSX("xlsx");

    private String value;

    WorkBookTypeEnum(String value){
        this.value = value;
    }

    public String getValue() {
        return value;
    }
}