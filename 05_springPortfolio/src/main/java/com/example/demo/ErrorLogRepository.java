package com.example.demo;

import org.springframework.data.jpa.repository.JpaRepository;
/**
 * CustomerエンティティのCRUD操作を可能とするインターフェース
 */
public interface ErrLogRepository extends JpaRepository<ErrLog, Long> {
}