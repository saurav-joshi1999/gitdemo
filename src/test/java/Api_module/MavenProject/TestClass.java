package Api_module.MavenProject;

import java.io.IOException;
import java.util.ArrayList;

public class TestClass {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		ExemlTest et = new ExemlTest();
		ArrayList<String> al = et.excelMethod("Purchase");
		for (String a : al) {
			System.out.println(a);
		}
		System.out.println(al.get(0));

	}

}
