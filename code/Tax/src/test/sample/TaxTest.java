package sample;
import static org.junit.Assert.*;
import org.junit.Test;

public class TaxTest {
	@Test
	public void test1() {
		assertEquals(Tax.calcConsumptionTax(100), 8);
	}
	
	@Test
	public void test2() {
		assertEquals(Tax.calcConsumptionTax(500), 40);
	}
	
	@Test
	public void test3() {
		assertEquals(Tax.calcConsumptionTax(1000), 80);
	}
	
	
}