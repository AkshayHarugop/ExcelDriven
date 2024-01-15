import java.io.IOException;
import java.util.ArrayList;

public class testSample {

	public static void main(String[] args) throws IOException {
		System.out.println("Start");
		dataDriven dD = new dataDriven();
		ArrayList<String> a = dD.getData("UserName","Book1");
		System.out.println(a);
//		for(int i=1;i<=a.size();i++) {
//			System.out.println(a.get(i-1));
//		}
		System.out.println("Done");
	}

}