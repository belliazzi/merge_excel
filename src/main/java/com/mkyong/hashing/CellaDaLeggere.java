package com.mkyong.hashing;

import org.apache.poi.ss.usermodel.CellType;

public class CellaDaLeggere {

	

	public CellaDaLeggere(int row, int cellNum, CellType type, String nome) {
		super();
		this.row = row;
		this.cellNum = cellNum;
		this.type = type;
		this.nome = nome;
	}

	public int getCellNum() {
		return cellNum;
	}

	public void setCellNum(int cellNum) {
		this.cellNum = cellNum;
	}

	int row;
	int  cellNum;
	org.apache.poi.ss.usermodel.CellType type;
    String nome;
	public int getRow() {
		return row;
	}

	public void setRow(int row) {
		this.row = row;
	}

	public org.apache.poi.ss.usermodel.CellType getType() {
		return type;
	}

	public void setType(org.apache.poi.ss.usermodel.CellType type) {
		this.type = type;
	}

	public String getNome() {
		return nome;
	}

	public void setNome(String nome) {
		this.nome = nome;
	}

	

}
