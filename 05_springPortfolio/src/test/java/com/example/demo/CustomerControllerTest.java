package com.example.demo;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.test.web.servlet.MockMvc;

@SpringBootTest
@AutoConfigureMockMvc 
public class CustomerControllerTest{
    /**
     * カスタマーコントローラー テストクラス
     */
    @Autowired
    private MockMvc mockMvc;
    @Autowired
    private JdbcTemplate jdbcTemplate;
    
    /**
     * データベースにデータを登録できること
     */
    @Test
    public void name16Test() throws Exception {
        mockMvc.perform(post("/customer")
        		.param("name", "JUnit 成功1")	//  
                .param("balance", "65536") )		// 文字列扱いだが念のため16進数の境界値のテストを自動化しておく
        		.andExpect(status().isOk()); 
    }
    /**
    * 金額フィールドに入力されていない場合、200を返し入力チェック後、データベースに登録しないこと
    */
   @Test
   public void balanceEmpTest() throws Exception {
	   String nameValue = "JUnit 失敗1";
	   
	   /** ヒット件数 カウント */
       Integer count = jdbcTemplate.queryForObject(
               "SELECT count(*) FROM customer WHERE cust_name = ?", Integer.class, nameValue);
       
	   /** 金額フィールド 空文字 送信  */
       mockMvc.perform(post("/customer")
       		.param("name",nameValue )		//  
               .param("balance", "") )			// 入力なし
       			.andExpect(status().isOk()); 
       
       /** ヒット件数 カウント  */
       Integer count2 = jdbcTemplate.queryForObject(
               "SELECT count(*) FROM customer WHERE cust_name = ?", Integer.class, nameValue);
      
       /** 件数比較  */
       assertEquals(count, count2, "異常値がDBに登録されました。【DBErr1】");
   }
    /**
     * 金額フィールドに０円が入力された場合、200を返し入力チェック後、データベースに登録しないこと
     */
    @Test
    public void balance0Test() throws Exception {
        mockMvc.perform(post("/customer")
        		.param("name", "JUnit 失敗2")	//  
                .param("balance", "0") )				// 0円
        		.andExpect(status().isOk()); 
    }
    /**
     * 金額フィールドに負数が入力された場合、200を返し入力チェック後、データベースに登録しないこと
     */
    @Test
    public void balanceMinusTest() throws Exception {
        mockMvc.perform(post("/customer")
        		.param("name", "JUnit 失敗3")	//  
                .param("balance", "-3") )			// 負数
        		.andExpect(status().isOk()); 
    }
    /**
     * 顧客名に記号が含まれていた場合、データベースに登録されないことを確認する
     */
    @Test
    public void nameEmpTest() throws Exception {
        mockMvc.perform(post("/customer")
        		.param("name", "")	 				// 未入力
                .param("balance", "08") )		// 文字列扱いだが念のため８進数のテストを自動化しておく
        		.andExpect(status().isOk()); 
    }
    /**
     * 顧客名に記号が含まれていた場合、データベースに登録されないことを確認する
     */
    @Test
    public void nameVal() throws Exception {
        mockMvc.perform(post("/customer")
        		.param("name", "性(旧性) 失敗4")	 // ()を含む形式
                .param("balance", "09") )				// 文字列扱いだが念のため８進数のテストを自動化しておく
        		.andExpect(status().isOk()); 
    }
    /**
     *  金額に文字列が入力された場合、レスポンス４００を返却すること
     */
    @Test
    public void balanceTypeTest() throws Exception {
        mockMvc.perform(post("/customer")
        		.param("name", "JUnit 失敗5")	// 
                .param("balance", "moji") )		// 文字列型
        		.andExpect(status().isBadRequest()); 
    }
}
