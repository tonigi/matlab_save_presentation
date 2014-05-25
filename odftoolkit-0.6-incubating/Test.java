/*
  cp=`echo *.jar | sed 's/ /:/g'`

  javac -cp odfdom-java-0.8.9-incubating.jar:simple-odf-0.8-incubating.jar:xslt-runner-1.2.2-incubating.jar:xslt-runner-task-1.2.2-incubating.jar Test.java 

  java -cp $cp:. Test

  java -cp $cp -jar /usr/share/java/js.jar
*/



import java.io.File;
import java.net.*;

import org.odftoolkit.simple.*;
import org.odftoolkit.simple.presentation.*;
import org.odftoolkit.simple.presentation.Slide.*;

/**
 * This class is a demo to show the presentation API of Simple Java API for ODF.
 * 
 *
 */
public class Test {

	public static void main(String[] args) throws Exception {
	    PresentationDocument document = PresentationDocument.newPresentationDocument();
	    int n=document.getSlideCount();
	    document.newSlide(n, "new slide", null);
	    URI imageuri=new URI("../x.png");
	    document.newImage(imageuri);
	    document.save("shit.odp");
	}
	
}

