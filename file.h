#pragma once
#include "libxl-4.0.1.0/include_cpp/libxl.h"
#include <iostream>


class File {

private:
	
	libxl::Book* book;
	libxl::Sheet* sheet;


public:

	libxl::Sheet* get_sheet() { 
		return sheet; 
	}

	libxl::Book* get_book() {
		return book;
	}
	void create_file() {
		book = xlCreateBook();
		
	}

	void create_sheet(const char* sheet_name) {
		sheet = book->addSheet(sheet_name);
		
	}

	void open_file(const char* file_name) {

		try {
			
			book->load(file_name);
			sheet = book->getSheet(0);
		}
		catch (const std::exception& e) {

			std::cout << "Can not open file: " << e.what() << std::endl;
		}
	}

	void save_file(const char* file_name) {
	
		book->save(file_name);
	}

	void print_sheet() {
	
		for (int row = sheet->firstRow(); row < sheet->lastRow(); ++row)
		{
			for (int col = sheet->firstCol(); col < sheet->lastCol(); ++col)
			{
				std::string str = sheet->readStr(row, col);
			
				std::cout << "(" << row << ", " << col << ") = " << str << std::endl;
			}
		}
	}

	void set_format_and_font(int size, const char* font_name) {

		libxl::Font* textFont = book->addFont();
		textFont->setSize(8);
		textFont->setName(font_name);
		libxl::Format* textFormat = book->addFormat();
		textFormat->setFont(textFont);
		
	}

	void write_string(int col, int row, const char* string) {
	
		sheet->writeStr(col, row, string);

	}

	const char* read_string(int col, int row) {

		const char* string = sheet->readStr(col, row);

		return string;
	}

	int first_column() {
	
		int first_col = sheet->firstCol();

		return first_col;
	}

	int last_column() {

		int last_col = sheet->lastCol();

		return last_col;
	}

	int first_row() {

		int first_row = sheet->firstRow();

		return first_row;
	}

	int last_row() {

		int last_row = sheet->lastRow();

		return last_row;
	}

};
