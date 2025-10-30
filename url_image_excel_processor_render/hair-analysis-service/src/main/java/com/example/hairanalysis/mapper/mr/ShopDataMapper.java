package com.example.hairanalysis.mapper.mr;

import com.example.hairanalysis.model.ShopData;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;
import org.apache.ibatis.annotations.Select;

@Mapper
public interface ShopDataMapper {

    @Select("SELECT id, name, num, pid, pname FROM shop_data WHERE id = #{id} AND status = 0")
    ShopData findById(@Param("id") String id);
}
