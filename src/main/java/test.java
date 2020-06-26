public class test {
    public static void main(String[] args) {
        String s = "1 724.10";
        String b = s.replaceAll("\\s+", "");
        System.out.println(b);
    }
}
