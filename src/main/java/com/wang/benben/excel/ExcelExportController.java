package com.wang.benben.excel;

import com.github.liaochong.myexcel.core.DefaultExcelBuilder;
import com.github.liaochong.myexcel.utils.AttachmentExportUtil;
import com.github.liaochong.myexcel.utils.FileExportUtil;
import com.wang.benben.excel.dto.ArtCrowd;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by zhongruo
 * Date: 20/3/5
 * Time: 上午11:54
 * Described:
 * @author zhongruo
 */
@Controller
public class ExcelExportController {

    @GetMapping("/default/excel/example")
    public void defaultBuild(HttpServletResponse response) throws Exception {
        List<ArtCrowd> dataList = this.getDataList();
        Workbook workbook = DefaultExcelBuilder.of(ArtCrowd.class)
                .build(dataList);
        AttachmentExportUtil.export(workbook, "艺术生信息", response);
    }

    @GetMapping("/default/excel/defaultBuildWithPassword")
    public void defaultBuildWithPassword(HttpServletResponse response) throws Exception {
        List<ArtCrowd> dataList = this.getDataList();
        Workbook workbook = DefaultExcelBuilder.of(ArtCrowd.class)
                .build(dataList);
        AttachmentExportUtil.encryptExport(workbook, "艺术生信息", response,"123456");
    }

    @GetMapping("/default/excel/defaultBuildFile")
    public void defaultBuildFile(HttpServletResponse response) throws Exception {
        List<ArtCrowd> dataList = this.getDataList();
        Workbook workbook = DefaultExcelBuilder.of(ArtCrowd.class)
                .build(dataList);
        FileExportUtil.export(workbook, new File("/Users/zhongruo/demo.xlsx"));
    }

    @GetMapping("/default/excel/demopwd")
    public void defaultBuildFileWithPwd() throws Exception {
        List<ArtCrowd> dataList = this.getDataList();
        Workbook workbook = DefaultExcelBuilder.of(ArtCrowd.class)
                .build(dataList);
        FileExportUtil.encryptExport(workbook, new File("/Users/zhongruo/demopwd.xlsx"),"123456");
    }

    private List<ArtCrowd> getDataList() {
        List<ArtCrowd> dataList = new ArrayList<>(100000);
        for (int i = 0; i < 100000; i++) {
            ArtCrowd artCrowd = new ArtCrowd();
            if (i % 2 == 0) {
                artCrowd.setName("Tom");
                artCrowd.setAge(19);
                artCrowd.setGender("Man");
                artCrowd.setPaintingLevel("一级证书");
                artCrowd.setDance(false);
                artCrowd.setAssessmentTime(LocalDateTime.now());
                artCrowd.setHobby(null);
            } else {
                artCrowd.setName("Marry");
                artCrowd.setAge(18);
                artCrowd.setGender("Woman");
                artCrowd.setPaintingLevel("一级证书");
                artCrowd.setDance(true);
                artCrowd.setAssessmentTime(LocalDateTime.now());
                artCrowd.setHobby("钓鱼");
            }
            dataList.add(artCrowd);
        }
        return dataList;
    }
}
