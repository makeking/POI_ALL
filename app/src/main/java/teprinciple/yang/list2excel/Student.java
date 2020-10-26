package teprinciple.yang.list2excel;

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
}
