package com.wang.benben.excel.dto;

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.annotation.ExcelModel;

import java.time.LocalDateTime;
import java.util.Date;

/**
 * Created by zhongruo
 * Date: 20/3/5
 * Time: 上午11:09
 * Described:
 */
@ExcelModel(sheetName = "艺术生", useFieldNameAsTitle = true)
public class ArtCrowd extends People{
    @ExcelColumn(order = 3, index = 3)
    private String paintingLevel;

    @ExcelColumn(order = 4, title = "是否会跳舞", groups = {People.class, String.class}, index = 4)
    private boolean dance;

    @ExcelColumn(order = 5, title = "考核时间", groups = {People.class, String.class}, index = 5)
    private LocalDateTime assessmentTime;

    @ExcelColumn(order = 6,index = 6,defaultValue = "---")
    private String hobby;

    public String getPaintingLevel() {
        return paintingLevel;
    }

    public void setPaintingLevel(String paintingLevel) {
        this.paintingLevel = paintingLevel;
    }

    public boolean isDance() {
        return dance;
    }

    public void setDance(boolean dance) {
        this.dance = dance;
    }

    public LocalDateTime getAssessmentTime() {
        return assessmentTime;
    }

    public void setAssessmentTime(LocalDateTime assessmentTime) {
        this.assessmentTime = assessmentTime;
    }

    public String getHobby() {
        return hobby;
    }

    public void setHobby(String hobby) {
        this.hobby = hobby;
    }

    @Override
    public String toString() {
        return "ArtCrowd{" +
                "paintingLevel='" + paintingLevel + '\'' +
                ", dance=" + dance +
                ", assessmentTime=" + assessmentTime +
                ", hobby='" + hobby + '\'' +
                '}';
    }
}
