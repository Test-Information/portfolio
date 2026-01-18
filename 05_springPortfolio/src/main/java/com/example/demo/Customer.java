package com.example.demo;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

/**
 * 顧客情報データベースエンティティ
 * 顧客作成画面から作成した顧客情報をデータベースに格納するモデル
 */
@Entity
@Table(name="customer")
@Data
public class Customer {
	/** 顧客ID */
    @Id	//主キー制約
    @GeneratedValue(strategy = GenerationType.IDENTITY)	// 自動シーケンス生成
    @Column(name="cust_id")
    private Long custId;
    
    /** 顧客名 */
    @Column(name="cust_name")
    private String custName;
    
    /** 残高 */
    @Column(name="cust_balance")
    private Long custBalance;
}