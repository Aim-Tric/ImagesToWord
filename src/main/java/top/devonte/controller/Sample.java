package top.devonte.controller;

import com.baidu.aip.ocr.AipOcr;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.util.*;

public class Sample {

    public static final String APP_ID = "";
    public static final String API_KEY = "";
    public static final String SECRET_KEY = "";

    public static final String PARSE_RESULT_KEY = "words_result";
    public static final String WORDS_KEY = "words";

    public static final String BASE_PATH = "sources";
    public static final String OUTPUT_PATH = "outputs";

    public static void main(String[] args) throws IOException, InterruptedException {
        AipOcr client = new AipOcr(APP_ID, API_KEY, SECRET_KEY);

        // 从当前路径获取doc，将doc或者image读取进来，
        File sources = new File(BASE_PATH);
        File output = new File(OUTPUT_PATH);

        List<String> fileNames = new ArrayList<>();

        Map<String, List<JSONObject>> resultMap = new HashMap<>();
        if (sources.exists()) {
            File[] files = sources.listFiles();
            if (files != null) {
                for (File file : files) {
                    String fileName = file.getName();
                    fileNames.add(fileName);
                }
            }

            HashMap<String, String> options = new HashMap<>();

            if (fileNames.size() > 0) {
                for (String name : fileNames) {
                    String relatePath = BASE_PATH + File.separator + name;
                    List<JSONObject> list = new ArrayList<>();
                    Thread.sleep(1000);
                    System.out.println("正在处理[" + relatePath + "]文件...");
                    if (name.endsWith(".docx")) {
                        XWPFDocument xwpfDocument = new XWPFDocument(new FileInputStream(new File(relatePath)));
                        List<XWPFPictureData> allPictures = xwpfDocument.getAllPictures();
                        for (XWPFPictureData picture : allPictures) {
                            byte[] data = picture.getData();
                            JSONObject jsonObject = client.basicAccurateGeneral(data, options);
                            list.add(jsonObject);
                        }
                    } else if (name.endsWith(".doc")) {
                        HWPFDocument hwpfDocument = new HWPFDocument(new FileInputStream(new File(relatePath)));
                        PicturesTable picturesTable = hwpfDocument.getPicturesTable();
                        List<Picture> allPictures = picturesTable.getAllPictures();
                        for (Picture picture : allPictures) {
                            byte[] content = picture.getContent();
                            JSONObject jsonObject = client.basicAccurateGeneral(content, options);
                            list.add(jsonObject);
                        }
                    } else if (name.endsWith(".png")) {
                        JSONObject jsonObject = client.basicAccurateGeneral(relatePath, options);
                        list.add(jsonObject);
                    }
                    resultMap.put(name, list);
                }
            }

            if (!output.exists()) {
                boolean mkdir = output.mkdir();
                if (!mkdir) {
                    System.out.println("创建输出文件夹失败，请手动创建output文件夹后再次运行程序。");
                    return;
                }
            }

            Set<String> keySet = resultMap.keySet();
            for (String key : keySet) {
                File file = new File(OUTPUT_PATH + File.separator + key + ".docx");
                XWPFDocument document = new XWPFDocument();
                FileOutputStream out = new FileOutputStream(file);
                List<JSONObject> jsonObjects = resultMap.get(key);
                for (JSONObject res : jsonObjects) {
                    JSONArray jsonArray = res.getJSONArray(PARSE_RESULT_KEY);
                    writeNewWords(document, jsonArray);
                }
                document.write(out);
                out.close();
                document.close();
            }
        } else {
            System.out.println("必须将要转换的图片或包含图片的word放置到sources文件夹下");
        }
    }

    public static void writeNewWords(XWPFDocument document, JSONArray jsonArray) {
        int row = jsonArray.length();
        XWPFParagraph paragraph = document.createParagraph();
        for (int i = 0; i < row; i++) {
            JSONObject rowObject = jsonArray.getJSONObject(i);
            String words = rowObject.getString(WORDS_KEY);
            words = words.replaceAll(",", "，");
            XWPFRun run = paragraph.createRun();
            run.setText(words);
        }
    }

}
