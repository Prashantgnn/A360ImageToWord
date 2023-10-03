package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;

import static com.automationanywhere.commandsdk.model.AttributeType.FILE;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

@BotCommand

//CommandPks adds required information to be dispalable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "ImageToWord", label = "[[Image To Word.label]]",
        node_label = "[[Image To Word.node_label]]", description = "[[Image To Word.description]]", icon = "pkg.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[Image To Word.return_label]]", return_type = STRING, return_description = "[[Image To Word.return_label_description]]")
public class ImageToWord {

    @Execute
    public StringValue action(
            //Idx 1 would be displayed first, with a text box for entering the value.
            @Idx(index = "1", type = FILE)
            //UI labels.
            @Pkg(label = "[[Image To Word.inputFilePath.label]]")
            //Force Word selection
            @FileExtension("jpeg,jpg,gif,png,webp")
            //Ensure that a validation error is thrown when the value is null.
            @NotEmpty
                    String sourceFolderPath,

            //Set Optional Export Dir
            @Idx(index = "2", type = FILE)
            @Pkg(label = "[[Image To Word.outputFilePath.label]]")
            @NotEmpty
                    String destinationFolderPath){
        if ("".equals(sourceFolderPath.trim()))
              throw new BotCommandException("Please select a valid file for processing.");
        if ("".equals(destinationFolderPath.trim()))
            throw new BotCommandException("Please select a valid file for processing.");

        // Get a list of image files from the source folder
        File sourceFolder = new File(sourceFolderPath);
        File[] imageFiles = sourceFolder.listFiles((dir, name) -> name.toLowerCase().endsWith(".jpg") ||
                name.toLowerCase().endsWith(".png") ||
                name.toLowerCase().endsWith(".gif") ||
                name.toLowerCase().endsWith(".webp") ||
                name.toLowerCase().endsWith(".jpeg"));

        if (imageFiles == null || imageFiles.length == 0) {
            //System.out.println("No image files found in the source folder.");
            return null;
        }

        // Create the destination folder if it doesn't exist
        File destinationFolder = new File(destinationFolderPath);
        if (!destinationFolder.exists()) {
            destinationFolder.mkdirs();
        }
        String outputFilePath = "";
        try {
            // Process each image file
            XWPFDocument document = new XWPFDocument();
            for (File imageFile : imageFiles) {

                // Initialize a new Word document
                //document = new XWPFDocument();
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();

                // Load the image and add it to the document
                FileInputStream imageStream = new FileInputStream(imageFile);
                run.addPicture(imageStream, XWPFDocument.PICTURE_TYPE_JPEG, imageFile.getName(), Units.toEMU(300), Units.toEMU(200));
                imageStream.close();
            }
            // Save the document to the destination folder
            outputFilePath = Paths.get(destinationFolderPath, "/convertedImageToWord.docx").toString();
            FileOutputStream out = new FileOutputStream(outputFilePath);
            document.write(out);
            out.close();
            document.close();

            //return new StringValue(outputFilePath);
        }
        catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }

        return new StringValue(outputFilePath);
    }
}




