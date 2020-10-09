package sample;
 
public class Tax {
	public String test;
 
	public static void main(String[] args) {
		int price = Integer.parseInt(args[0]);
		int consumptionTax = calcConsumptionTax(price);
		System.out.println(consumptionTax);
	}

	public static int calcConsumptionTax(int price) {
		int res = (int)(price * 0.08);
		if(res <= 10) {
			System.out.println("A");
		} else if(res >= 100) {
			System.out.println("B");
		} else {
			System.out.println("C");
		}
		return res;
	}
}