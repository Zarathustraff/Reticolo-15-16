// AnalisiDati.cpp : definisce il punto di ingresso dell'applicazione console.
//

#include "stdafx.h"
#include "BasicExcel.hpp"
#include "ExcelFormat.h"

#ifdef _WIN32

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shellapi.h>
#include <crtdbg.h>

#else // _WIN32

#define	FW_NORMAL	400
#define	FW_BOLD		700

#endif // _WIN32



int main()
{


	YExcel::BasicExcel Excel("Excel.xls"); /*Prova, Serve a vedere se funziona qualcosa*/

	/*Excel.New(1);

	Excel.SaveAs("Excel.xls");*/

	YExcel::BasicExcelWorksheet* sheet = Excel.GetWorksheet(0);

	YExcel::BasicExcelCell* cell = sheet->Cell(0, 0);

	double val;

	cell->Get(val);

	std::cout << val << endl;

	int a;

	std::cout << "Inserire intero." << endl;

	std::cin >> a;

    return 0;
}

