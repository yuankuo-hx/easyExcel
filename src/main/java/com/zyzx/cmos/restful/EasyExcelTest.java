package com.zyzx.cmos.restful;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.metadata.Sheet;
import com.zyzx.cmos.entity.Student;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;
import org.springframework.web.multipart.MultipartRequest;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("test")
public class EasyExcelTest {
    private final Logger logger = LoggerFactory.getLogger(EasyExcelTest.class);

    /**
     * 本地Excel读取
     *
     */
    @RequestMapping("read")
    public void read() {
        logger.info("read 读取文件开始。。。。");
        String filePath = "D:/IDEA wordspace/Excel/学生表.xlsx";
        File file = new File(filePath);
        List<Object> list = new ArrayList<>();
        try {
            list = readWithoutHead(new FileInputStream(file));
        } catch (Exception e) {
            e.printStackTrace();
        }
        list.remove(0);
        for (int i = 0; i < list.size(); i++) {
            Student student = (Student) list.get(i);
            logger.info("第" + i + "个学生信息： "
                    + " 学号：" + student.getStudentId()
                    + " 姓名：" + student.getName()
                    + " 性别：" + student.getSex());
        }
        logger.info("read 读取文件结束。。。。");
    }

    /**
     * 数据写入本地Excel
     *
     */
    @RequestMapping("write")
    public void write() {
        logger.info("write 写入文件开始。。。。");
        List<Object> objects = new ArrayList<>();
        objects.add(new Student("a", "110", "幻想", "男"));
        objects.add(new Student("b", "119", "幻灵", "男"));
        objects.add(new Student("c", "120", "幻月", "女"));
        writeWithoutHead("D:/IDEA wordspace/Excel/新学生表.xlsx", objects);
        logger.info("write 写入文件结束。。。。");

    }

    /**
     * 请求Excel读取
     *
     */
    @RequestMapping("readUploadFile")
    public void readUploadFile(MultipartRequest request) {
        logger.info("readUploadFile 读取文件开始。。。。");
        if (request instanceof MultipartHttpServletRequest) {
            MultipartHttpServletRequest multipartRequest = (MultipartHttpServletRequest) request;
            // 通过表单中的参数名来接收文件流（可用 file.getInputStream() 来接收输入流）
            MultipartFile file = multipartRequest.getFile("file");
            logger.info("上传的文件名称:" + file.getOriginalFilename());
            logger.info("上传的文件大小:" + file.getSize());
            // 接收其他表单参数
            String usrId = multipartRequest.getParameter("usrId");
            logger.info("usrId:" + usrId);

            List<Object> studentList = new ArrayList<>();
            try {
                studentList = readWithoutHead(file.getInputStream());
            } catch (Exception e) {
                logger.error("表格读取失败");
            }
            // 对表格中的数据作处理
            studentList.remove(0);
            for (int i = 0; i < studentList.size(); i++) {
                Student student = (Student) studentList.get(i);
                logger.info("第" + i + "个学生信息： "
                        + " 学号：" + student.getStudentId()
                        + " 姓名：" + student.getName()
                        + " 性别：" + student.getSex());
            }

        }
        logger.info("readUploadFile 读取文件结束。。。。");
    }


    public List<Object> readWithoutHead(InputStream in) {
        List<Object> read = EasyExcel.read(in, new Sheet(0, 0, Student.class));
        return read;
    }

    public void writeWithoutHead(String fileName, List<Object> objects) {
        EasyExcel.write(fileName, Student.class).sheet("学生列表").doWrite(objects);
    }
}
