package model;

/**
 * @program: ExcelUtil
 * @description:
 * @author: zan.Kang
 * @create: 2019-06-04 08:38
 **/
public class Student {

    private String name;
    public Integer age;
    private String school;

    public void setName(String name) {
        this.name = name;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public void setSchool(String school) {
        this.school = school;
    }

    public String getName() {
        return name;
    }

    public Integer getAge() {
        return age;
    }

    public String getSchool() {
        return school;
    }

    @Override
    public String toString() {
        return "[name:" + name + ",age:" + age + ",school:" + school + "]";
    }
}
