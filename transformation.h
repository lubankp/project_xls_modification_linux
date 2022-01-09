#pragma once
#include "file.h"

class Transformation {

private:
	File file1;
	File file2;

public:

	void modification() {
	
		for (int row = file1.first_row(); row < file1.last_row(); ++row)
		{
			for (int col = file1.first_column(); col < file1.last_column(); ++col)
			{
				
				file2.write_string(col, row, file1.read_string(row,col));
			}
		}
	
	}

	void transform_file(const wchar_t* file_name1, const wchar_t* file_name2) {
	
		file1.create_file();
		file2.create_file();

		file2.set_format_and_font(8, L"Century Gothic");
		file2.create_sheet(L"New");

		file1.open_file(file_name1);
		modification();
		//file1.print_sheet();
		file2.save_file(file_name2);
	}

};