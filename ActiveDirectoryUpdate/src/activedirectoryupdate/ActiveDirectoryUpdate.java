package activedirectoryupdate;

import java.io.File;
import java.io.FileReader;
import java.io.BufferedReader;
import java.io.FilenameFilter;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.util.Date;
import java.text.SimpleDateFormat;

/**
 * ActiveDirectoryUpdate.java - This program massively executes an Active Directory update<br/>
 * Usage - <i>java ActiveDirectoryUpdate [Folder Name]</i><br/>
 * Output - <i>ActiveDirectoryUpdate.bat</i><br/>
 * @author Glenn Dimaliwat 12/04/2015
 */
public class ActiveDirectoryUpdate {
    
    /**
     * Calls the main program for ActiveDirectoryUpdate
     * @param args Folder Name
     * @throws Exception 
     */
    public static void main(String[] args) throws Exception {
        
        // Validate input folder
        if(args.length == 0) {
            System.out.println("Proper Usage: \"java ActiveDirectoryUpdate [Folder Name]\"");
            System.exit(0);
        }
        else {
            
            // Reads the input folder for files
            readFiles(args[0]);

            // Clean up batch folders
            cleanup();

            // Garbage collect
            System.gc();
        }
    }
    
    /**
     * Scans the folder and reads all the files
     * @param folderName Contains all the input files
     * @throws Exception 
     */
    private static void readFiles(String folderName) throws Exception {
        
        // Initialize batch number
        int batch = 1;

        // Create Log Folder
        File logFolder = new File(".\\Logs");
        if (!logFolder.exists()) {
            logFolder.mkdir();
        }

        // Define Input Folder
        File folder = new File(folderName);
        File[] listOfFiles = folder.listFiles();

        for(File f : listOfFiles) {
            if (f.isFile()) {
                String batchFileName = generateBatchFile(folderName+"\\"+f.getName(),batch);

                System.out.println("Processing..."+folderName+"\\"+f.getName());
                // Run the batch job
                executeBatchFile(batchFileName);
                
                // Archive the File
                archiveFile(folderName+"\\"+f.getName());

                // Next batch
                batch++;
            }
        }
    }
    
    /**
     * Generates a Batch File
     * @param inputFile Input File Name
     * @param batchNumber Batch Number of this Run
     * @return outputFile
     * @throws Exception 
     */
    private static String generateBatchFile(String inputFile,int batchNumber) throws Exception {
        
        final int AD_ACCOUNT_COLUMN = 0;
        final int FIRST_NAME_COLUMN = 1;
        final int LAST_NAME_COLUMN = 2;
        final int DISPLAY_NAME_COLUMN = 3;
        final int DEPARTMENT_COLUMN = 4;
        final int LEVEL_COLUMN = 5;
        final int VIP_COLUMN = 6;
        final int PHONE_NUMBER_COLUMN = 7;

        String outputFile = "File"+batchNumber+".bat";
        PrintWriter pWriter;
        try (BufferedReader fh = new BufferedReader(new FileReader(inputFile))) {
            pWriter = new PrintWriter(outputFile);
            String s;
            int x=0;
            while ((s=fh.readLine())!=null) {
                if (x>0) { // Do not include header
                    String f[] = s.split("\\|");
                    String AD_ACCOUNT = f[AD_ACCOUNT_COLUMN];
                    String FIRST_NAME = f[FIRST_NAME_COLUMN];
                    String LAST_NAME = f[LAST_NAME_COLUMN];
                    String DISPLAY_NAME = f[DISPLAY_NAME_COLUMN];
                    String DEPARTMENT = f[DEPARTMENT_COLUMN];
                    String LEVEL = f[LEVEL_COLUMN];
                    String VIP = f[VIP_COLUMN];
                    String PHONE_NUMBER = f[PHONE_NUMBER_COLUMN];

                    String cmd = "powershell.exe -Command \"Set-ADUser "+AD_ACCOUNT+" -Replace @{GivenName='"+FIRST_NAME+"';DisplayName='"+DISPLAY_NAME+"';Department='"+DEPARTMENT+"';Title='"+LEVEL+"';Comment='"+VIP+"';TelephoneNumber='"+PHONE_NUMBER+"'} -Surname '"+LAST_NAME+"'\"";

                    pWriter.println(cmd);
                }
                x = x + 1;
            }
            pWriter.close();
        }
        
        return outputFile;
    }

    /**
     * Executes a batch run
     * @param batchFilename Batch File to be processed
     * @throws Exception 
     */
    private static void executeBatchFile(String batchFilename) throws Exception {

        Date date = new Date() ;
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss-SSS") ;
        
        try (PrintWriter logWriter = new PrintWriter(".\\Logs\\ActiveDirectory-Log "+dateFormat.format(date)+".txt")) {
            String line;

            // Executable batch file
            String command = "\".\\"+batchFilename+"\"";

            // Executing the command
            Process powerShellProcess = Runtime.getRuntime().exec(command);

            // Getting the results
            powerShellProcess.getOutputStream().close();

            // Begin Output
            System.out.println("Standard Output:");
            logWriter.println("Standard Output:");
            try (BufferedReader stdout = new BufferedReader(new InputStreamReader(
                 powerShellProcess.getInputStream()))) {
                while ((line = stdout.readLine()) != null) {
                    System.out.println(line);
                    logWriter.println(line);
                }
                stdout.close();
            }

            // Output Errors
            System.out.println("\r\nStandard Error:");
            logWriter.println("\r\nStandard Error:");
            try (BufferedReader stderr = new BufferedReader(new InputStreamReader(
                 powerShellProcess.getErrorStream()))) {
                while ((line = stderr.readLine()) != null) {
                    System.out.println(line);
                    logWriter.println(line);
                }
                stderr.close();
            }
            logWriter.close();
        }

        System.out.println("\r\n\r\nUpdate Done!");
    }
    
    /**
     * Archive the input file to Archive Folder
     * @param filename Filename to be archived
     * @throws Exception 
     */
    private static void archiveFile(String filename) throws Exception {
        
        // Move file command
        String command = "cmd /c move "+filename+" Archive";
        
        // Executing the command
        Process archiveProcess = Runtime.getRuntime().exec(command);
        archiveProcess.getOutputStream().close();
    }
    
    /**
     * Removes temporary batch files
     * @throws Exception 
     */
    private static void cleanup() throws Exception {
        
        // Define Input Folder
        File folder = new File(".");
        File[] listOfFiles = folder.listFiles( new FilenameFilter() {
            @Override
            public boolean accept( final File dir,
                                   final String name ) {
                return name.matches("File.*\\.bat");
            }
        } );
        
        for(File f : listOfFiles) {
            if (f.isFile()) {
                f.delete();
            }
        }
    }
}