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

double convertiSecondiInGradi(double secondi);

double modulo(double val);

double convertiSecondiInGradi(double secondi) {

	int Gradi = secondi;

	double Gr = Gradi;

	double finale = ((secondi - Gr) / 0.6) + Gr;

	return finale;

};

double modulo(double val) {
	if (val < 0.0) {
		return -val;
	}
	else {
		return val;
	};
}




int main()
{


	YExcel::BasicExcel Excel("Excel.xls"); /*Prova, Serve a vedere se funziona qualcosa*/

	/*Excel.New(1);

	Excel.SaveAs("Excel.xls");*/

	const double thetagrande = convertiSecondiInGradi(269.50); //in gradi

	const double thetapiccolo = convertiSecondiInGradi(90.00); //in gradi

	std::cout << "thetagrande: " << thetagrande << ", thetapiccolo: " << thetapiccolo << endl;

	/* YExcel::BasicExcelWorksheet* bluDx = Excel.GetWorksheet(0); 

	YExcel::BasicExcelWorksheet* azzurroDx = Excel.GetWorksheet(1);

	YExcel::BasicExcelWorksheet* verdeDx = Excel.GetWorksheet(2);

	YExcel::BasicExcelWorksheet* rossoDx = Excel.GetWorksheet(3);

	YExcel::BasicExcelWorksheet* bluSx = Excel.GetWorksheet(4);

	YExcel::BasicExcelWorksheet* azzurroSx = Excel.GetWorksheet(5);

	YExcel::BasicExcelWorksheet* verdeSx = Excel.GetWorksheet(6);

	YExcel::BasicExcelWorksheet* rossoSx = Excel.GetWorksheet(7); */

	int col=0, row=0, number=0;

	double thetap, thetam;

	for (number = 0; number < 8; number++) {

		for (row = 0; !(row>4||(row>3&&(number==3||number==7))); row++) {

			YExcel::BasicExcelWorksheet* sheet = Excel.GetWorksheet(number);

			YExcel::BasicExcelCell* cellA = sheet->Cell(row, col);

			YExcel::BasicExcelCell* cellB = sheet->Cell(row, col + 2);

			YExcel::BasicExcelCell* errorCell = sheet->Cell(row, col + 1);

			std::cout << "row: " << row << ", col: " << col << ", number: " << number << endl;
			
			std::cout << "CellA: " << cellA->GetDouble() << endl << "CellB: " << cellB->GetDouble() << endl;

			thetap = convertiSecondiInGradi(cellA->GetDouble());
			thetam = convertiSecondiInGradi(cellB->GetDouble());

			std::cout << "thetap: " << thetap << endl << "thetam: " << thetam << endl;

			double set;
			set = (modulo(thetap - thetagrande) + modulo(thetam - thetapiccolo))*0.5;
			cellA->SetDouble(set);
			errorCell->SetInteger(row+1);
			cellB->EraseContents();
			std::cout <<"set: "<< set << endl;
		};

	};

	for (number = 8; number < 9; number++) { //l'ultima scheda è per la prima parte, ovvero il calcolo dell'errore statistico.

		for (row = 0; row < 10; row++) {

			YExcel::BasicExcelWorksheet* sheet = Excel.GetWorksheet(number);

			YExcel::BasicExcelCell* cellA = sheet->Cell(row, col);

			/*YExcel::BasicExcelCell* cellB = sheet->Cell(row, col + 2);*/

			std::cout << "row: " << row << ", col: " << col << ", number: " << number << endl;

			std::cout << "CellA: " << cellA->GetDouble() << endl/* << "CellB: " << cellB->GetDouble() << endl*/;

			thetap = convertiSecondiInGradi(cellA->GetDouble());
			//thetam = convertiSecondiInGradi(cellB->GetDouble());

			std::cout << "thetap: " << thetap << endl /*<< "thetam: " << thetam << endl*/;

			double set;
			set = modulo(thetap - thetagrande);
			cellA->SetDouble(set);
			//cellB->EraseContents();
			std::cout << "set: " << set << endl;

		}; //qua finisce il calcolo/conversione dei dati per l'errore statistico.

	};
	/* YExcel::BasicExcelCell* cell1 = sheet->Cell(0, 0);

	YExcel::BasicExcelCell* cell2 = sheet->Cell(0, 1);

	double val;

	val = cell1->GetDouble();

	double finale = convertiSecondiInGradi(val);

	cell2->SetDouble(finale);

	std::cout << val << endl;

	std::cout << "il valore" << val << " convertito è: " << finale << endl; */

	int a; 

	std::cout << "Inserire intero. Programma Terminato." << endl;

	std::cin >> a;

	Excel.SaveAs("Exceloutput.xls");

    return 0;
}

