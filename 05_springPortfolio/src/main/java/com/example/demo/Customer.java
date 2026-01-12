package com.example.demo;

import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;

/**
 * 顧客情報データベースエンティティ
 * 顧客作成画面から作成した顧客情報をデータベースに格納するモデル
 */
@Entity
public class Customer {
    @Id	//主キー制約
    @GeneratedValue(strategy = GenerationType.IDENTITY)	// 作成時の自動シーケンス生成
    private Long id;				// 顧客ID
    private String name;		// 顧客名
    private Long balance;	// 残高

    // Getter・Setter
    public Long getId() { return id; }
    public void setId(Long id) { this.id = id; }
    public String getName() { return name; }
    public void setName(String name) { this.name = name; }
    public Long getBalance() { return balance; }
    public void setBalance(Long balance) { this.balance = balance; }
}