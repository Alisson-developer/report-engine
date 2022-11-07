package reportengine.data;

import reportengine.annotations.ColumnReport;
import reportengine.annotations.Report;

@Report(name = "Relatorio-anotado", sheetName = "Aba 1")
public class ObjectDataDTO {

    @ColumnReport(title = "Nome do Cidadão")
    private String name;
    private int age;
    private float height;

    public ObjectDataDTO() {
    }

    public ObjectDataDTO(String name, int age, float height) {
        this.name = name;
        this.age = age;
        this.height = height;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public float getHeight() {
        return height;
    }

    public void setHeight(float height) {
        this.height = height;
    }

    @Override
    public String toString() {
        return "ObjectDataDTO {" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", height=" + height +
                '}';
    }
}
