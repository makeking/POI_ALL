package teprinciple.yang.list2excel.bean;

/**
 * Created by Administrator on 2017/4/13.
 */

public class Student {

    String name;
    String sex;
    int age;
    int id;
    String classNo;
    int math;
    int english;
    int chinese;

    public Student(String name, String sex, int age, int id, String classNo, int math, int english, int chinese) {
        this.name = name;
        this.sex = sex;
        this.age = age;
        this.id = id;
        this.classNo = classNo;
        this.math = math;
        this.english = english;
        this.chinese = chinese;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getClassNo() {
        return classNo;
    }

    public void setClassNo(String classNo) {
        this.classNo = classNo;
    }

    public int getMath() {
        return math;
    }

    public void setMath(int math) {
        this.math = math;
    }

    public int getEnglish() {
        return english;
    }

    public void setEnglish(int english) {
        this.english = english;
    }

    public int getChinese() {
        return chinese;
    }

    public void setChinese(int chinese) {
        this.chinese = chinese;
    }
}
