package com.example.demo;

import java.time.LocalDateTime;

import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;

/**
 * エラーログ エンティティ
 * カスタムエラー画面遷移時の情報を格納するモデル
 */
@Entity
public class ErrLog {
    @Id	//主キー制約
    @GeneratedValue(strategy = GenerationType.IDENTITY)	// 作成時の自動シーケンス生成
    private long errId;			// エラーID
    private LocalDateTime eDate;	// エラー時刻
    private int status;		// Httpステータスコード
    private String path;		// エラー発生URI

    // Getter・Setter
    public long getErrId() { return errId; }
    public void setErrId(Long errId) { this.errId = errId; }
    public LocalDateTime getEDate() { return eDate; }
    public void setEDate(	LocalDateTime eDate) { this.eDate = eDate; }
    public int getStatus() { return status; }
    public void setStatus(int status) { this.status = status; }
    public String getPath() { return path; }
    public void setPath(String path) { this.path = path; }
}