package Main;

import java.util.ArrayList;
import java.util.List;

import org.testng.TestNG;

public class Main_Run  {

    public static void main(String[] args) {
        
        // TODO Auto-generated method stub
        // Create object of TestNG Class
        TestNG runner=new TestNG();

        // Create a list of String 
        List<String> suitefiles=new ArrayList<String>();

        // Add xml file which you have to execute
        suitefiles.add("E:\\selenium\\SEO_content\\practice\\testng.xml");

        // now set xml file for execution
        runner.setTestSuites(suitefiles);

        // finally execute the runner using run method
        runner.run();
        }
    }

