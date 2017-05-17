package sdsreport;

/**
 *
 * @author Glenn Dimaliwat
 */
public final class Utilities {
    
    protected static String fixString(String str) {
        if(str!=null) {
            str = str.replaceAll("null", "").replaceAll("__b", " ").replaceAll("__7", "&").replaceAll("__u", "-").replaceAll("__P", "(").replaceAll("__p", ")").replaceAll("__M", ",").replaceAll("__f","/").replaceAll("__A", "+").replaceAll("__a", "'").replaceAll("__d", ".");
        }
        else {
            str = "";
        }
        return str;
    }
    
    protected static String unfixString(String str) {
        if(str!=null) {
            // Note: Plus/Parenthesis/Dot signs could not be converted back due to warnings - Glenn
            str = str.replaceAll(" ", "__b").replaceAll("&", "__7").replaceAll("-", "__u").replaceAll(",", "__M").replaceAll("/","__f").replaceAll("'", "__a");
        }
        else {
            str = "";
        }
        return str;
    }
    
    protected static String removeOrganization(String str) {
        String strOrig = str;
        if(str!=null) {
            str = str.replaceAll("Asset Management", "").replaceAll("Business Analyst", "").replaceAll("Developer", "").replaceAll("Field Support", "").replaceAll("Network and Communications", "").replaceAll("Operations Management", "").replaceAll("SMO", "").replaceAll("Service Desk", "").replaceAll("System Administrator", "");
            str = str.trim().replaceAll("  ", " ");
            
            if(str.equalsIgnoreCase("")) {
                str = strOrig;
            }
        }
        else {
            str = "";
        }
        return str;
    }
}
