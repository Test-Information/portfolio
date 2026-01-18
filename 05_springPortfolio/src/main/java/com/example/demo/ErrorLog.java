package com.example.demo;

import java.time.LocalDateTime;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

/**
 * エラーログ エンティティ
 * カスタムエラー画面遷移時の情報を格納するモデル
 */
@Entity
@Data
@Table(name="err_log")
public class ErrorLog {
	/** エラーID */
    @Id	//主キー制約
    @GeneratedValue(strategy = GenerationType.IDENTITY)	// 自動シーケンス生成
    @Column(name="err_id")
    private long errId;
    
    /** エラー時刻 */
    @Column(name="err_date")
    private LocalDateTime errDate;
    
	/** Httpステータスコード */
    @Column(name="err_status")
    private int errStatus;
    
	/** エラー発生URI */
    @Column(name="err_path")
    private String errPath;
}