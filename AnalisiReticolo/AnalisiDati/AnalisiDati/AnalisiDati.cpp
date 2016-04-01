// AnalisiDati.cpp : definisce il punto di ingresso dell'applicazione console.
//

#include "stdafx.h"
#include "BasicExcel.hpp"
#include "ExcelFormat.h"
#include "math.h"

#ifdef _WIN32

#define WIN32_LEAN_AND_MEAN


#include <windows.h>
#include <shellapi.h>
#include <crtdbg.h>

#else // _WIN32


#define	FW_NORMAL	400
#define	FW_BOLD		700

#endif // _WIN32
double sind(double gradi);

double cosd(double gradi);

double convertiSecondiInGradi(double secondi);

double modulo(double val);

double gradiRadianti(double gradi);

double convertiSecondiInGradi(double secondi) {

	int Gradi = secondi;

	double Gr = Gradi;

	double finale = ((secondi - Gr) / 0.6) + Gr;

	return finale;

};

double gradiRadianti(double gradi) {
	const double M_PI = 4 * atan(1);
	return gradi*M_PI / 180;
};

double sind(double gradi) {
	const double M_PI = 4 * atan(1);
	return sin((gradi) * M_PI / 180);
};

double cosd(double gradi) {
	const double M_PI = 4 * atan(1);
	return cos((gradi) * M_PI / 180);
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
	const double d = 12.65e-6;

	const double dErr = 0.05e-6;

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

	double thetap=0, thetam=0;

	for (number = 0; number < 8; number++) {

		for (row = 0; !(row>4||(row>3&&(number==3||number==7))); row++) {

			YExcel::BasicExcelWorksheet* sheet = Excel.GetWorksheet(number);

			YExcel::BasicExcelCell* cellA = sheet->Cell(row, col);

			std::cout << "cellA output: " << cellA->GetDouble() << endl;

			YExcel::BasicExcelCell* cellB = sheet->Cell(row, col + 2);

			YExcel::BasicExcelCell* cellC = sheet->Cell(row, col + 1);

			

			std::cout << "row: " << row << ", col: " << col << ", number: " << number << endl;

			double doubleAa = cellA->GetDouble();
			double doubleBb = cellB->GetDouble();
			
			thetap = convertiSecondiInGradi(doubleAa);
			thetam = convertiSecondiInGradi(doubleBb);

			std::cout << "thetap: " << thetap << endl << "thetam: " << thetam << endl;

			double set;
			set = (modulo(thetap - thetagrande) + modulo(thetam - thetapiccolo))*0.5;
			double dSin;
			dSin = d*sind(set);
			double errorDSin = sqrt(((sind(set)*dErr)*(sind(set)*dErr)) + ((d*cosd(set)*gradiRadianti(0.03))*(d*cosd(set)*gradiRadianti(0.03))));
			cellC->SetDouble(set);
			cellA->SetInteger(row+1);
			cellB->SetDouble(0.03);
			sheet->Cell(row, col + 3)->SetDouble(dSin);
			sheet->Cell(row, col + 4)->SetDouble(errorDSin);
			std::cout <<"set: "<< set <<", d*sin(theta): "<< dSin << endl;
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

