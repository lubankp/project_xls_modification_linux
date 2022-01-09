#include <iostream>
#include "file.h"
#include "transformation.h"

int main() {

	Transformation transform;
	transform.transform_file(L"file.xls", L"file1.xls");
	
}
