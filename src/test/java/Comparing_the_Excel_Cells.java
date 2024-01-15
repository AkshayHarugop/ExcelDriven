import java.io.IOException;
import java.util.ArrayList;

public class Comparing_the_Excel_Cells {
	public static void main(String[] args) throws IOException {
		dataDriven1 DD = new dataDriven1();
		ArrayList<String> a = DD.getData("Book1", "testdata", "Testcases");
//		System.out.println(a);
		ArrayList<String> b = DD.getData("Book2", "testdata", "Testcases");
//		System.out.println(b);
		
		System.out.println(a.containsAll(b));
	}
}
 