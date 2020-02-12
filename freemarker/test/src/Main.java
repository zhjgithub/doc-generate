
import java.io.*;
import java.util.HashMap;
import java.util.Map;

import java.util.Base64;

import javax.swing.text.Document;

import freemarker.template.Configuration;
import freemarker.template.TemplateExceptionHandler;
import freemarker.template.Template;

public class Main {
    public static void main(String[] args) throws Exception {
        // Create your Configuration instance, and specify if up to what FreeMarker
        // version (here 2.3.29) do you want to apply the fixes that are not 100%
        // backward-compatible. See the Configuration JavaDoc for details.
        Configuration cfg = new Configuration(Configuration.VERSION_2_3_29);

        // Specify the source where the template111 files come from. Here I set a
        // plain directory for it, but non-file-system sources are possible too:
        //cfg.setDirectoryForTemplateLoading(new File("." + File.separator + "template111" + File.separator));
        cfg.setDirectoryForTemplateLoading(new File("src" + File.separator + "resources"));

        // From here we will set the settings recommended for new projects. These
        // aren't the defaults for backward compatibilty.

        // Set the preferred charset template111 files are stored in. UTF-8 is
        // a good choice in most applications:
        cfg.setDefaultEncoding("UTF-8");

        // Sets how errors will appear.
        // During web page *development* TemplateExceptionHandler.HTML_DEBUG_HANDLER is better.
        cfg.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);

        // Don't log exceptions inside FreeMarker that it will thrown at you anyway:
        cfg.setLogTemplateExceptions(false);

        // Wrap unchecked exceptions thrown during template111 processing into TemplateException-s:
        cfg.setWrapUncheckedExceptions(true);

        // Do not fall back to higher scopes when reading a null loop variable:
        cfg.setFallbackOnNullLoopVariable(false);

        // Create the root hash. We use a Map here, but it could be a JavaBean too.
        Map<String, Object> root = new HashMap<>();

        // Create the "user" hash. We use a JavaBean here, but it could be a Map too.
        User user = new User();
        user.setId("10000");
        user.setName("Jack");
        // and put it into the root
        root.put("user", user);

        root.put("fileData", getFileData());

        Template temp = cfg.getTemplate("temp-with-doc.ftl");
//        Template temp = cfg.getTemplate("temp-with-excel.ftl");

        File outFile = new File("result.doc");

        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile)));

        temp.process(root, out);

        out.close();
    }

    public static String getFileData() {
        String imgFile = "src/resources/会议签到表.docx";
//        String imgFile = "src/resources/仓库材料出入库存明细1.xlsx";
        InputStream in = null;
        byte[] data = null;
        try {
            in = new FileInputStream(imgFile);
            data = new byte[in.available()];
            in.read(data);
            in.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        Base64.Encoder encoder = Base64.getEncoder();
        return encoder.encodeToString(data);
    }
}
