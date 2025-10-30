package com.example.hairanalysis.config;

import javax.sql.DataSource;

import org.apache.ibatis.session.SqlSessionFactory;
import org.mybatis.spring.SqlSessionFactoryBean;
import org.mybatis.spring.SqlSessionTemplate;
import org.mybatis.spring.annotation.MapperScan;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.boot.autoconfigure.jdbc.DataSourceProperties;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.core.io.support.PathMatchingResourcePatternResolver;

@Configuration
@MapperScan(basePackages = "com.example.hairanalysis.mapper.boka", sqlSessionTemplateRef = "bokaSqlSessionTemplate")
public class BokaDatabaseConfig {

    @Bean
    @ConfigurationProperties("spring.datasource.boka")
    public DataSourceProperties bokaDataSourceProperties() {
        return new DataSourceProperties();
    }

    @Bean
    public DataSource bokaDataSource(@Qualifier("bokaDataSourceProperties") DataSourceProperties properties) {
        return properties.initializeDataSourceBuilder().build();
    }

    @Bean
    public SqlSessionFactory bokaSqlSessionFactory(@Qualifier("bokaDataSource") DataSource dataSource) throws Exception {
        SqlSessionFactoryBean factoryBean = new SqlSessionFactoryBean();
        factoryBean.setDataSource(dataSource);
        factoryBean.setMapperLocations(new PathMatchingResourcePatternResolver()
                .getResources("classpath*:mapper/boka/*Mapper.xml"));
        return factoryBean.getObject();
    }

    @Bean
    public SqlSessionTemplate bokaSqlSessionTemplate(@Qualifier("bokaSqlSessionFactory") SqlSessionFactory sqlSessionFactory) {
        return new SqlSessionTemplate(sqlSessionFactory);
    }
}
