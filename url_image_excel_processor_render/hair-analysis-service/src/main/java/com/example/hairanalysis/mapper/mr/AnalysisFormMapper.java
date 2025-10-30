package com.example.hairanalysis.mapper.mr;

import com.example.hairanalysis.model.AnalysisResult;
import org.apache.ibatis.annotations.Insert;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

@Mapper
public interface AnalysisFormMapper {

    @Insert("INSERT INTO analysis_from (dep_name, stroe_id, mf_ty_count, mf_ty_amount, mf_cj_count, mf_cj_amount, dt_ty_count, dt_ty_amount, dt_cj_count, dt_cj_amount, data_date, create_time, hz_ty_count, hz_ty_amount, hz_cj_count, hz_cj_amount, hz_cj_ratio, store_name) " +
            "VALUES (#{result.departmentName}, #{result.storeId}, #{result.hairExperienceCount}, #{result.hairExperienceAmount}, #{result.hairDealCount}, #{result.hairDealAmount}, #{result.groundPushExperienceCount}, #{result.groundPushExperienceAmount}, #{result.groundPushDealCount}, #{result.groundPushDealAmount}, #{result.month}, NOW(), #{result.totalExperienceCount}, #{result.totalExperienceAmount}, #{result.totalDealCount}, #{result.totalDealAmount}, 0, #{result.storeName})")
    void insert(@Param("result") AnalysisResult result);
}
