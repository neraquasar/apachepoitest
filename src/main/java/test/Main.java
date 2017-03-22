package test;

/**
 * Main class to test files creating
 *
 * @author Konstantin Valerievich Dichenko
 * @version 1.0
 */
public class Main
{
    static DocService docService = new DocService();

    public static void main(String[] args)
    {
        docService.saveIn("c:/users/kd/desktop/Test-Doc.docx");

    }
}
