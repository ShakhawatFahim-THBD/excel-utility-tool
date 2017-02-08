package net.therap.util;

/**
 * @author shakhawat.hossain
 * @since 2/8/17
 */
public class StringUtil {

    public static boolean isEmpty(String val) {
        return val == null || val.trim().isEmpty();
    }

    public static boolean isNotEmpty(String val) {
        return !isEmpty(val);
    }
}
