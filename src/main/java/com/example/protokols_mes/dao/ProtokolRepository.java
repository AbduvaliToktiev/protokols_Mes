package com.example.protokols_mes.dao;

import com.example.protokols_mes.entity.Protokol;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface ProtokolRepository extends JpaRepository<Protokol, Long> {

    @Query(value = "select * from mes_protokol.protokols p " +
            "where p.name like %:keyword%", nativeQuery = true)
    List<Protokol> findByKeyword(@Param("keyword") String keyword);
}
