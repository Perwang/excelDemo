package com.wang.benben.excel.service;

import com.github.liaochong.myexcel.core.DefaultExcelReader;
import com.wang.benben.excel.dto.ArtCrowd;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

/**
 * Created by zhongruo
 * Date: 20/3/5
 * Time: 上午11:05
 * Described:
 */
@RunWith(SpringRunner.class)
@SpringBootTest
public class DefaultExcelImportTest {

    @Test
    public void importALLTest(){
        URL htmlToExcelEampleURL = this.getClass().getResource("/templates/read_example.xlsx");
        try {

            Path path = Paths.get(htmlToExcelEampleURL.toURI());
            long startTime=System.currentTimeMillis();
            // 方式一：全部读取后处理
            List<ArtCrowd> result = DefaultExcelReader.of(ArtCrowd.class)
                    .sheet(0) // 0代表第一个，如果为0，可省略该操作，也可sheet("名称")读取
                    .rowFilter(row -> row.getRowNum() > 0) // 如无需过滤，可省略该操作，0代表第一行
                    .read(path.toFile());// 可接收inputStream
            System.out.println("读取时间："+(System.currentTimeMillis()-startTime));
            result.parallelStream().forEach(t->{
                System.out.println(t.toString());
            });
            System.out.println(System.currentTimeMillis()-startTime);
        } catch (URISyntaxException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void importByOneTest(){
        URL htmlToExcelEampleURL = this.getClass().getResource("/templates/read_example.xlsx");
        Path path = null;
        try {
            path = Paths.get(htmlToExcelEampleURL.toURI());
            long startTime=System.currentTimeMillis();
            // 方式二：读取一行处理一行，可自行决定终止条件
// readThen有两种重写接口，返回Boolean型接口允许在返回False情况下直接终止读取
            DefaultExcelReader.of(ArtCrowd.class)
                    .sheet(0) // 0代表第一个，如果为0，可省略该操作，也可sheet("名称")读取
                    .rowFilter(row -> row.getRowNum() > 0) // 如无需过滤，可省略该操作，0代表第一行
                    //.beanFilter(ArtCrowd::isDance) // bean过滤
                    .readThen(path.toFile() ,artCrowd -> {System.out.println(artCrowd.toString());});// 可接收inputStream
            System.out.println(System.currentTimeMillis()-startTime);
        } catch (URISyntaxException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }


    }
}
