package test;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.function.Predicate;
import java.util.stream.Collectors;

/**
 * @author zhangpengyue
 * @date 2018/12/9
 */
public class Apple {
    private String color;

    public Apple(String color) {
        this.color = color;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public boolean isGreen(Apple apple) {
        return this.color.equals("green");
    }

    public double weight(){
        return Math.random();
    }

    @Override
    public String toString() {
        return "Apple{" +
                "color='" + color + '\'' +
                '}';
    }


    public static void main(String[] args) {
        List<Apple> al = new ArrayList<>();
        al.add(new Apple("green"));
        al.add(new Apple("red"));
        List<Apple> result = al.stream().filter((apple) -> apple.getColor().equals("green")).collect(Collectors.toCollection(ArrayList::new));
        System.out.println(result);

    }
}
