package test;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Class for different actions with prepared MS Word files
 *
 * @author Konstantin Valerievich Dichenko
 * @version 1.0
 */
public class DocService
{
    private DocConstructor docConstructor = new DocConstructor();

    public void saveIn(String path)
    {
        try
        {
            FileOutputStream fileOutputStream = new FileOutputStream(path);
            docConstructor.getDoc().write(fileOutputStream);
            System.out.println("File was written");
            fileOutputStream.close();
        } catch (IOException e)
        {
            System.out.println(e.getMessage());
        }
    }
}
