import com.automationanywhere.botcommand.ImageToWord;
import com.automationanywhere.botcommand.data.impl.StringValue;
import org.testng.Assert;
import org.testng.annotations.Test;

public class TestClass {
    @Test
    public void testIMAGEtoWORD(){
        String inputFile = "D:\\inputfile";
        String outputPath = "E:\\outfile";

        ImageToWord imageToword = new ImageToWord();

        StringValue outputFile = imageToword.action(inputFile,outputPath);
        Assert.assertEquals(outputFile.toString(), "E:\\outfile\\convertedImageToWord.docx");
    }
}
