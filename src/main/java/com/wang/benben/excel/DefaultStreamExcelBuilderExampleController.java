package com.wang.benben.excel;

import com.github.liaochong.myexcel.core.DefaultStreamExcelBuilder;
import com.github.liaochong.myexcel.core.strategy.WidthStrategy;
import com.github.liaochong.myexcel.utils.AttachmentExportUtil;
import com.wang.benben.excel.dto.ArtCrowd;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.Executors;

/**
 * Created by zhongruo
 * Date: 20/3/5
 * Time: 下午4:42
 * Described:
 */
@Controller
public class DefaultStreamExcelBuilderExampleController {

    /**
     * 导出单个sheet
     * @param response
     * @throws Exception
     */
    @GetMapping("/default/excel/stream/example")
    public void streamBuild(HttpServletResponse response) throws Exception {
        try (DefaultStreamExcelBuilder<ArtCrowd> defaultExcelBuilder = DefaultStreamExcelBuilder.of(ArtCrowd.class)
                .widthStrategy(WidthStrategy.CUSTOM_WIDTH)
                .hasStyle()
                .threadPool(Executors.newFixedThreadPool(10))
                .start()) {
            List<CompletableFuture> futures = new ArrayList<>();
            for (int i = 0; i < 100; i++) {
                CompletableFuture future = CompletableFuture.runAsync(() -> {
                    List<ArtCrowd> dataList = this.getDataList();
                    defaultExcelBuilder.append(dataList);
                });
                futures.add(future);
            }
            futures.forEach(CompletableFuture::join);
            Workbook workbook = defaultExcelBuilder.build();
            AttachmentExportUtil.export(workbook, "艺术生信息1", response);
        }
    }

    /**
     * 设置每个excel的大小
     * @param response
     * @throws Exception
     */
    @GetMapping("/default/excel/stream/examplegz")
    public void examplegz(HttpServletResponse response) throws Exception {
        try (DefaultStreamExcelBuilder<ArtCrowd> defaultExcelBuilder = DefaultStreamExcelBuilder
                .of(ArtCrowd.class)
                .threadPool(Executors.newFixedThreadPool(10))
                .capacity(10_000)
                .start()) {
            List<CompletableFuture> futures = new ArrayList<>();
            for (int i = 0; i < 100; i++) {
                CompletableFuture future = CompletableFuture.runAsync(() -> {
                    List<ArtCrowd> dataList = this.getDataList();
                    defaultExcelBuilder.append(dataList);
                });
                futures.add(future);
            }
            futures.forEach(CompletableFuture::join);
            // 最终构建
            Path zip = defaultExcelBuilder.buildAsZip("test");
            AttachmentExportUtil.export(zip,"finalName.zip",response);
        }
    }


    /**
     * 导出多个sheet内容
     * @param response
     * @throws Exception
     */
    @GetMapping("/default/excel/stream/continue/example")
    public void streamBuildWithContinue(HttpServletResponse response) throws Exception {
        DefaultStreamExcelBuilder<ArtCrowd> defaultExcelBuilder = DefaultStreamExcelBuilder.of(ArtCrowd.class)
                .threadPool(Executors.newFixedThreadPool(10))
                .start();

        List<CompletableFuture> futures = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            CompletableFuture future = CompletableFuture.runAsync(() -> {
                List<ArtCrowd> dataList = this.getDataList();
                defaultExcelBuilder.append(dataList);
            });
            futures.add(future);
        }
        futures.forEach(CompletableFuture::join);
        Workbook workbook = defaultExcelBuilder.build();

        DefaultStreamExcelBuilder<ArtCrowd> defaultStreamExcelBuilder = DefaultStreamExcelBuilder.of(ArtCrowd.class, workbook)
                .threadPool(Executors.newFixedThreadPool(10))
                .sheetName("sheet2")
                .start();
        futures = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            CompletableFuture future = CompletableFuture.runAsync(() -> {
                List<ArtCrowd> dataList = this.getDataList();
                defaultStreamExcelBuilder.append(dataList);
            });
            futures.add(future);
        }
        futures.forEach(CompletableFuture::join);
        workbook = defaultStreamExcelBuilder.build();
        AttachmentExportUtil.export(workbook, "艺术生信息", response);
    }

    private List<ArtCrowd> getDataList() {
        List<ArtCrowd> dataList = new ArrayList<>(1000);
        for (int i = 0; i < 1000; i++) {
            ArtCrowd artCrowd = new ArtCrowd();
            if (i % 2 == 0) {
                artCrowd.setName("Tom");
                artCrowd.setAge(19);
                artCrowd.setGender("Man");
                artCrowd.setPaintingLevel("一级证书");
                artCrowd.setDance(false);
                artCrowd.setAssessmentTime(LocalDateTime.now());
                artCrowd.setHobby("摔跤");
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
