import com.google.common.base.CharMatcher;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * @author: YunTao.Li
 * @create: 2020-05-06 23:07
 * @description:
 **/
public class StringFinderMain {

    // 查找关键字
    public static final String CONDITIONSTR = "普元信息技术";

    // 查找目录
    public static final String WORKDIR = "D:\\归档\\2020.04.02_cmmi3\\resource";

    public List<String> fileLog = new ArrayList<String>();
    public List<String> errorLog = new ArrayList<String>();

    public static void main(String[] args) {
        try {
            StringFinderMain finder = new StringFinderMain();
            finder.searchDir(WORKDIR, true);

            for (int i = 0; i < finder.fileLog.size(); i++) {
                String log = finder.fileLog.get(i);
                System.out.println(log);
            }
            System.out.println("======================================================================================================================================");
            System.out.println("======================================================================================================================================");
            System.out.println("======================================================================================================================================");
            System.out.println("======================================================================================================================================");
            System.out.println("======================================================================================================================================");
            System.out.println("======================================================================================================================================");

            for (int i = 0; i < finder.errorLog.size(); i++) {
                String err = finder.errorLog.get(i);
                System.out.println(err);
            }
        } catch (Throwable e) {
            e.printStackTrace();
        }
    }

    /**
     * @param dir     : 查找路径,记住这里应该传一个路径而不是文件
     * @param recurse : 是否递归查找
     * @return : void
     * @author : YunTao.Li
     * @date : 2020/4/29 2020/4/29
     */
    public void searchDir(String dir, boolean recurse) throws Throwable {

        try {
            // step1 : 判断是否是一个目录，不是目录直接返回
            File dirFile = new File(dir);
            if (!dirFile.isDirectory()) {
                return;
            }

            // step2 : 获得目录下的所有文件，并遍历这些文件（或者目录）
            File[] files = dirFile.listFiles();

            for (int i = 0; i < files.length; i++) {

                try {
                    // step2.1 : 如果遍历到目录则递归查询子目录
                    if (recurse && files[i].isDirectory()) {
//                        System.out.println("===查找目录:"+files[i].getAbsolutePath());
                        searchDir(files[i].getAbsolutePath(), true);
                    } else {

                        // step2.1.1 : 获得当前文件名
                        String filename = files[i].getAbsolutePath();
//                        this.printMessage("===开始读文件：" + filename);

                        // step2.1.2 : 判断文件如果是压缩型的文件，则遍历压缩包内的所有文件
                        if (filename.endsWith(".jar") || filename.endsWith(".zip")) {
                            readZipFileLine(filename);
                        }
                        // step2.1.3 : 如果是word文件
                        else if (filename.endsWith(".doc") || filename.endsWith(".docx")) {
                            readDocFileLine(filename);
                        }else if (filename.endsWith(".xlsx") || filename.endsWith(".xls")) {
//                            readExcelFileLine(filename);
                            //TODO: read excel
                        } else if (filename.endsWith(".txt") || filename.endsWith(".md") || filename.endsWith(".js") || filename.endsWith(".json")) {
                            readSimpleFileLine(filename);
                        }
                        // step 2.1.4 : 如果是其他文件
                        else {
//                            this.printMessage("***跳过文件：" + filename);
                        }
                    }

                } catch (Throwable e1) {
                    errorLog.add(e1.getMessage());
                }
            }
        } catch (Throwable e) {
            e.printStackTrace();
            throw e;
        }
    }

    /**
     * 读普通文件，输出结果
     *
     * @param filePathName : 文件path名
     * @return : void
     * @author : YunTao.Li
     * @date : 2020/4/30 2020/4/30
     */
    public void readSimpleFileLine(String filePathName) throws Throwable {
        try {
            BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(filePathName)));
            while (reader.read() != -1) {
                String tempStr = reader.readLine();

                if (null != tempStr && tempStr.indexOf(CONDITIONSTR) > -1) {
                    this.printMessage("=" + filePathName + "文件中发现关键字  --->  " + CONDITIONSTR);
                    break;
                }
            }
        } catch (Throwable e) {
            throw new Throwable(filePathName + "-" + e.getMessage());
        }
    }

    /**
     * 读zip或者jar包文件，输出结果
     *
     * @param filePathName : 文件path名
     * @return : void
     * @author : YunTao.Li
     * @date : 2020/4/30 2020/4/30
     */
    public void readZipFileLine(String filePathName) throws Throwable {
        try {

            // step1 : 解压缩
            ZipFile zip = new ZipFile(filePathName);
            Enumeration entries = zip.entries();

            // step2 : 遍历压缩包里的每个文件
            while (entries.hasMoreElements()) {
                ZipEntry entry = (ZipEntry) entries.nextElement();
                StringBuffer className = new StringBuffer(entry.getName().replace("/", "."));
                String thisClassName = className.toString();

                BufferedReader reader = new BufferedReader(new InputStreamReader(zip.getInputStream(entry)));
                while (reader.read() != -1) {
                    String tempStr = reader.readLine();

                    if (null != tempStr && tempStr.indexOf(CONDITIONSTR) > -1) {
                        this.printMessage("=" + filePathName + "包里的" + thisClassName + "文件中发现关键字  --->  " + CONDITIONSTR);
                        break;
                    }
                }
            }


        } catch (Throwable e) {
            throw new Throwable(filePathName + "-" + e.getMessage());
        }
    }

    /**
     * 读DOC文件，输出结果
     *
     * @param filePathName : 文件path名
     * @return : void
     * @author : YunTao.Li
     * @date : 2020/4/30 2020/4/30
     */
    public void readDocFileLine(String filePathName) throws Throwable {
        InputStream stream = null;
        List<String> contextList = new ArrayList();
        try {
            stream = new FileInputStream(new File(filePathName));
            if (filePathName.endsWith(".doc")) {
                HWPFDocument document = new HWPFDocument(stream);
                WordExtractor extractor = new WordExtractor(document);
                String[] contextArray = extractor.getParagraphText();
                Arrays.asList(contextArray).forEach(context -> contextList.add(CharMatcher.whitespace().removeFrom(context)));
                extractor.close();
                document.close();
            } else if (filePathName.endsWith(".docx")) {
                XWPFDocument document = new XWPFDocument(stream).getXWPFDocument();
                List<XWPFParagraph> paragraphList = document.getParagraphs();
                paragraphList.forEach(paragraph -> contextList.add(CharMatcher.whitespace().removeFrom(paragraph.getParagraphText())));
                document.close();
            } else {
                System.out.println("不是word");
            }
        } catch (Throwable e) {
            throw new Throwable(filePathName + "-" + e.getMessage());
        } finally {
            if (null != stream) try {
                stream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        for (int i = 0; i < contextList.size(); i++) {
            String line = contextList.get(i);
            if (null != line && line.indexOf(CONDITIONSTR) > -1) {
                this.printMessage("=" + filePathName + "文件中发现关键字  --->  " + CONDITIONSTR);
                break;
            }
        }
    }

    public void readExcelFileLine() {

    }

    public void printMessage(String msg) {
        System.out.println(msg);
    }
}
