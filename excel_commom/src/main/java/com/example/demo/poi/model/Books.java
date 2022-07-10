package com.example.demo.poi.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class Books {
    private Integer id;

    private String bookname;

    private Double prices;

    private Integer counts;

    private int typeid;


	public Books(Integer id, String bookname) {
		super();
		this.id = id;
		this.bookname = bookname;
	}

    
    
}