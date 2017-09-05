import javax.swing.*;

public class JavaToExcelMain {
	   // Class constructor creates the frame and sets the frame main operations. 
		public static void main(String[] args) {
			JFrame frame = new JFrame("New Submission");
			frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
			
			ButtonDemoPanel2 panel = new ButtonDemoPanel2();
			frame.add(panel);
	

			frame.setSize(905, 600);
			frame.setResizable(false);
			frame.setVisible(true);
		}


}