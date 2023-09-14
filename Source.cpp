#define CURL_STATICLIB
#define _CRT_SECURE_NO_WARNINGS
#pragma once
#include <vector>
#include <string>
#include <iostream>
#include "curl\curl.h"
#include <windows.h>
#include <fstream>
#include <cmath>f
#include <OpenXLSX\OpenXLSX.hpp>
#include <stdlib.h>
#include "windows.h"


using namespace OpenXLSX;
using namespace std;

class money {
public:
	double thickness = 0;
	vector <string> vectorCutStringUcoin;
	string infoUcoin;
	string name;
	string name2;
	string token;
	int year;
	int number;
	string condition;
	double weight = 0;
	double diametr = 0;
	int edition = 0;
	int middlePrice1 = 0;
	int middlePrice2 = 0;
	double middlePrice3 = 0;
	double priceTraidUNCMax = 0;
	double priceTraidUNCMin = 0;
	double priceTraidCondMax = 0;
	double priceTraidCondMin = 0;
	string material = "";
	string coin_edge = "";
public:
	string url;
	string urlRaritetus;
	string urlCoinsmart;
	string htmlRaritetus;
	string htmlCoinsmart;
	string html3;
	money(int num) { number = num; };
	money() {};
	void GetMoneyFromExcel();
	void PostMoneyIntoExcel();
	void ParseUrlMoneyCoinsmart();
	void ParseUrlMoneyRaritetus();
	void MiddlePriseRaritetus();
	void GetInfoFromRaritetus();
	void MiddlePriseCoinsmart();
	void GetInfoFromCoinsmart();
	void GetWeightDiameterThicknessUcoin();
	void GetPriceEdition();
	void CutStringAllMoney(vector <money> Vector, int num, int k, string allInformation);
	void GetInfoFromThreeSite();
	void CutStringFromUcoin();
	void GetPriceUcoin();
	void GetMaterialUcoin();
};
size_t write_data(void* ptr, size_t size, size_t nmemb, std::string* data) {
	data->append((char*)ptr, size * nmemb);
	return size * nmemb;
}

string UTF8_to_CP1251(std::string const& utf8)
{
	if (!utf8.empty())
	{
		int wchlen = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), utf8.size(), NULL, 0);
		if (wchlen > 0 && wchlen != 0xFFFD)
		{
			std::vector<wchar_t> wbuf(wchlen);
			MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), utf8.size(), &wbuf[0], wchlen);
			std::vector<char> buf(wchlen);
			WideCharToMultiByte(1251, 0, &wbuf[0], wchlen, &buf[0], wchlen, 0, 0);

			return std::string(&buf[0], wchlen);
		}
	}
	return std::string();
}

//извлечение чисел в вектор из строки
template <typename T>
vector<T> num_from_string(string str, T x)
{
	vector <T> Tvector;
	char temp[1024];
	strcpy(temp, str.c_str());
	for (auto i = strtok(temp, " \f\n\r\t\v<>"""); i != nullptr; i = strtok(nullptr, " \f\n\r\t\v<>""")) {
		char* it;
		double num = strtod(i, &it);

		if (*it == '\0') {
			Tvector.push_back(num);
		}
	}
	return Tvector;
}

// получение html по ссылке
string GetDataFromSite(string url) {
	CURL* curl;
	CURLcode response;
	std::string data = "";
	curl = curl_easy_init();
	if (curl)
	{
		curl_easy_setopt(curl, CURLOPT_URL, url.c_str());
		curl_easy_setopt(curl, CURLOPT_WRITEFUNCTION, write_data);
		curl_easy_setopt(curl, CURLOPT_WRITEDATA, &data);
		response = curl_easy_perform(curl);
		if (response != CURLE_OK) {
			std::cerr << "Error: " << curl_easy_strerror(response) << std::endl;
			return "";
		}
		//else std::cout << data << std::endl;

		curl_easy_cleanup(curl);
	}
	return data;
};

//получение количества монет в файле
int count_money()
{
	XLDocument doc;
	doc.open("./111.xlsx");
	auto wks = doc.workbook().worksheet("Main");
	int B1 = wks.cell("B1").value();
	doc.close();
	return B1;
}

//получение информации о монете по порядковому номеру
void money::GetMoneyFromExcel()
{
	string num = to_string(number + 2);
	XLDocument doc;
	doc.open("./111.xlsx");
	auto wks = doc.workbook().worksheet("Main");
	string C = wks.cell("C" + num).value();
	int D = wks.cell("D" + num).value();
	string D2 = to_string(D);
	string F = wks.cell("F" + num).value();
	year = D;

	url = C + " " + D2 + " " + F;
	name = url;

	while (url.find(" ") != string::npos) {
		url.replace(url.find(" "), 1, "+");
	}
	string E = wks.cell("E" + num).value();
	condition = E;
	string G = wks.cell("G" + num).value();
	token = G;
	doc.close();
}

//запись даннных о монете
void money::PostMoneyIntoExcel()
{
	string num = to_string(number + 2);
	XLDocument doc;
	doc.open("./111.xlsx");
	auto wks = doc.workbook().worksheet("Main");
	wks.cell("T" + num).value() = middlePrice1;
	wks.cell("U" + num).value() = middlePrice2;
	wks.cell("V" + num).value() = middlePrice3;

	wks.cell("W" + num).value() = priceTraidUNCMax;
	wks.cell("X" + num).value() = priceTraidUNCMin;
	wks.cell("Y" + num).value() = priceTraidCondMax;
	wks.cell("Z" + num).value() = priceTraidCondMin;

	wks.cell("K" + num).value() = weight;
	wks.cell("N" + num).value() = diametr;
	wks.cell("L" + num).value() = edition;
	wks.cell("M" + num).value() = thickness;

	//wks.cell("AA" + num).value() = material;
	//wks.cell("AB" + num).value() = coin_edge;

	doc.save();
	doc.close();
}

//получение адреса страницы монеты из страницы поиска
void money::ParseUrlMoneyRaritetus()
{
	int StUrl, EnUrl;
	StUrl = htmlRaritetus.find("url='/stoimost-monet") + 5;
	EnUrl = htmlRaritetus.find("' class=");
	if (StUrl == 4 && EnUrl == -1)
		urlRaritetus = "";
	else {
		urlRaritetus = htmlRaritetus.substr(StUrl, EnUrl - StUrl);
	}
}

//получение средней цены по коду страницы и сохранности
void money::MiddlePriseRaritetus() {
	int StPrices, numbCond;
	string condidion_mas[] = { "G","VG","F","VF","XF","AU","UNC" };
	for (int i = 0; i < 7; i++)
		if (condidion_mas[i] == condition) {
			numbCond = i;
			break;
		}
	StPrices = htmlRaritetus.find("avg-prices");
	string Prices = htmlRaritetus.substr(StPrices + 106, 94);

	while (Prices.find("-") != std::string::npos) {
		Prices.replace(Prices.find("-"), 1, "0");
	}
	while (Prices.find(" ") != std::string::npos) {
		Prices.erase(Prices.find(" "), 1);
	}
	vector<int> condition_vec = num_from_string(Prices, numbCond);
	middlePrice1 = condition_vec[numbCond];
}

//получение веса, диаметра и тиража монеты по 1 сайту
void money::GetInfoFromRaritetus()
{
	int MoneyStr = htmlRaritetus.find("col-sm-4 descfullcont");

	string edition_s = htmlRaritetus.substr(MoneyStr - 80, 35);
	while (edition_s.find(" ") != std::string::npos) {
		edition_s.erase(edition_s.find(" "), 1);
	}
	int i = 1;
	vector<int> vec1 = num_from_string(edition_s, i);
	if (!vec1.empty() && edition == 0)
		edition = vec1[0];

	string weight_s = htmlRaritetus.substr(MoneyStr + 230, 150);
	while (weight_s.find(",") != std::string::npos) {
		weight_s.replace(weight_s.find(","), 1, ".");
	}
	double d = 0.1;
	vector<double> vec2 = num_from_string(weight_s, d);
	if (vec2.size() >= 2 && (weight == 0 || diametr == 0)) {
		weight = vec2[0];
		diametr = vec2[1];
	}

	if (weight == vec2[0] && diametr == vec2[1]) {
		this->MiddlePriseRaritetus();
	}
}


//-----------------------------------------//
//Функции для 2 сайта

//получение адреса страницы монеты из страницы поиска
void money::ParseUrlMoneyCoinsmart()
{
	int StUrl, EnUrl;
	string NewUrl;
	StUrl = htmlCoinsmart.find("data-url=") + 10;
	EnUrl = htmlCoinsmart.find("?cart=");
	if (StUrl == 9 && EnUrl == -1)
		urlCoinsmart = "";
	else {
		urlCoinsmart = htmlCoinsmart.substr(StUrl, EnUrl - StUrl);
	}
}

//получение цены с аукциона
void money::MiddlePriseCoinsmart()
{
	int priceMoneyStr = htmlCoinsmart.find("data-price=");
	string priceMoney = htmlCoinsmart.substr(priceMoneyStr, 50);

	while (priceMoney.find(" ") != std::string::npos) {
		priceMoney.erase(priceMoney.find(" "), 1);
	}

	int d = 1;
	vector<int> vec = num_from_string(priceMoney, d);

	middlePrice2 = vec[0];
}

//получение диаметра, веса и тиража
void money::GetInfoFromCoinsmart()
{
	int weightMoneyStr = htmlCoinsmart.find("features striped");
	string weight_s = "";
	if (weightMoneyStr != string::npos)
		weight_s = htmlCoinsmart.substr(weightMoneyStr + 250, 400);
	while (weight_s.find(",") != std::string::npos) {
		weight_s.replace(weight_s.find(","), 1, ".");
	}

	while (weight_s.find(" ") != std::string::npos) {
		weight_s.erase(weight_s.find(" "), 1);
	}

	double d = 0.1;
	vector<double> vec = num_from_string(weight_s, d);

	if (vec.size() >= 4 && weight == 0 && edition == 0 && diametr == 0)
	{
		weight = vec[1];
		edition = vec[3];
		diametr = vec[0];

	}
	if (vec.size() >= 4 && (weight == 0 || edition == 0 || diametr == 0) && (weight == vec[1] || edition == vec[3] || diametr == vec[0]))
	{
		if (weight == 0)
			weight = vec[1];
		if (edition == 0)
			edition = vec[3];
		if (diametr == 0)
			diametr = vec[0];
	}
	if (vec.size() >= 2 && weight == vec[1] && diametr == vec[0]) {
		this->MiddlePriseCoinsmart();
	}
}


void money::GetPriceEdition()
{
	try
	{
		string text = vectorCutStringUcoin[2];
		int positionEdition = text.find("Тираж");
		int positionPrice = text.find("Цена");


		if (positionEdition != string::npos && positionPrice != string::npos)
		{
			int positionYear = text.find(to_string(year));

			string rowsWithYear = text.substr(positionYear, text.length() - positionYear);

			string rowWithYear = rowsWithYear.substr(0, rowsWithYear.find("\n"));

			double j = 1;
			vector<double> vec2 = num_from_string(rowWithYear, j);
			this->middlePrice3 = vec2[vec2.size() - 1];

			int pos = rowWithYear.find(".");
			while (pos != string::npos)
			{
				rowWithYear.replace(pos, 1, "");
				pos = rowWithYear.find(".");
			}

			int i = 1;
			vector<int> vec1 = num_from_string(rowWithYear, i);
			this->edition = vec1[1];
		}
		else if (positionEdition != string::npos && positionPrice == string::npos)
		{
			string textCopy = text;
			int pos = textCopy.find(".");
			while (pos != string::npos)
			{
				textCopy.replace(pos, 1, "");
				pos = textCopy.find(".");
			}
			int i = 1;
			vector<int> vec1 = num_from_string(textCopy, i);
			this->edition = vec1[0];
		}
		else if (positionEdition == string::npos && positionPrice != string::npos)
		{
			int positionToken = text.find(this->token);
			if (positionToken != string::npos)
			{
				string rowsWithToken = text.substr(positionToken, text.length() - positionToken);
				string rowWithToken = rowsWithToken.substr(0, rowsWithToken.find("\n"));
				double i = 1.0;
				vector<double> vec1 = num_from_string(rowWithToken, i);
				this->middlePrice3 = vec1[0];
			}

			string textCopy = vectorCutStringUcoin[4];
			int pos = textCopy.find(".");
			while (pos != string::npos)
			{
				textCopy.replace(pos, 1, "");
				pos = textCopy.find(".");
			}
			int i = 1;
			vector<int> vec1 = num_from_string(textCopy, i);
			this->edition = vec1[0];
		}
	}
	catch (exception e) { std::cout << "Ошибка: " << e.what() << endl; }
}


void money::GetPriceUcoin()
{
	string text = vectorCutStringUcoin[3];
	string condidion_mas[] = { "G","VG","F","VF","XF","AU","UNC" };
	vector <string> vectorUcoin;
	int masPositionCondition[7];
	while (text.length() > 3)
	{
		int minPosition = 100000000;
		for (int i = 0; i < 7; i++)
		{
			masPositionCondition[i] = text.find(condidion_mas[i], 4);
		}
		for (int i = 0; i < 7; i++)
		{
			if ((masPositionCondition[i] < minPosition) && (masPositionCondition[i] > 0))
			{
				minPosition = masPositionCondition[i];
			}
		}
		if (minPosition < string::npos)
		{
			vectorUcoin.push_back(text.substr(0, minPosition));
			text.erase(0, minPosition);
		}
		else
		{
			vectorUcoin.push_back(text.substr(0, text.length() - 1));
			text.erase(0, text.length() - 1);
		}

	}
	for (int i = 0; i < vectorUcoin.size(); i++)
	{
		int posUNC = vectorUcoin[i].find("UNC");
		int posCond = vectorUcoin[i].find(condition);
		int posYear = vectorUcoin[i].find(to_string(year));
		if ((posUNC != -1) && (posUNC != string::npos) && (posYear != -1) && (posYear != string::npos))
		{
			string temp = vectorUcoin[i].substr(posYear + 4, vectorUcoin[i].length() - posYear);
			double d = 0.1;
			vector<double> tempV = num_from_string(temp, d);
			if (tempV.size() > 0)
			{
				if ((priceTraidUNCMax == 0) || (tempV[0] > priceTraidUNCMax))
				{
					priceTraidUNCMax = tempV[0];
				}
				if ((priceTraidUNCMin == 0) || (tempV[0] < priceTraidUNCMin))
				{
					priceTraidUNCMin = tempV[0];
				}
			}
		}
		if (condition != "UNC")
		{
			if ((posCond != -1) && (posCond != string::npos) && (posYear != -1) && (posYear != string::npos))
			{
				string temp = vectorUcoin[i].substr(posYear + 4, vectorUcoin[i].length() - posYear);
				double d = 0.1;
				vector<double> tempV = num_from_string(temp, d);
				if (tempV.size() > 0)
				{
					if ((priceTraidCondMax == 0) || (tempV[0] > priceTraidCondMax))
					{
						priceTraidCondMax = tempV[0];
					}
					if ((priceTraidCondMin == 0) || (tempV[0] < priceTraidCondMin))
					{
						priceTraidCondMin = tempV[0];
					}
				}
			}
		}
		else
		{
			priceTraidCondMax = priceTraidUNCMax;
			priceTraidCondMin = priceTraidUNCMin;
		}
	}
}

//функция для поиска материала и гурта
void money::GetMaterialUcoin() {
	try {
		string text = vectorCutStringUcoin[1];
		cout << text;
		int position1 = text.find("Материал: ");
		int position2 = text.find("Гурт: ");
		int position3 = text.find("Форма: ");
		material = text.substr(position1 + 10, position2 - position1 - 10);
		coin_edge = text.substr(position2 + 6, position3 - position2 - 7);
	}
	catch (...){
		cout << "Нет информации по материалу и гурту\n";
		material = "-";
		coin_edge = "-";
	}
}

void money::GetWeightDiameterThicknessUcoin() {
	try {
		string text = vectorCutStringUcoin[1];
		int position1 = text.find("Вес");
		string numText = text.substr(position1, text.length() - position1);
		double i = 1.1;
		vector<double> vec = num_from_string(numText, i);
		weight = vec[0];
		diametr = vec[1];
		thickness = vec[2];
	}
	catch (...) {}
}

void money::CutStringFromUcoin() {
	string infoUcoinText = this->infoUcoin;
	for (int i = 0; i < infoUcoinText.length(); i++) {
		if (infoUcoinText[i] == '\n' && infoUcoinText[i + 1] == '\n')
		{
			infoUcoinText[i + 1] = '№';
		}
	}
	while (infoUcoinText.length() > 0) {
		vectorCutStringUcoin.push_back(infoUcoinText.substr(0, infoUcoinText.find('№')));
		infoUcoinText.erase(0, infoUcoinText.find('№') + 1);
	}
}

void money::CutStringAllMoney(vector <money> Vector, int num, int k, string allInformation) {
	int positionBegin = allInformation.find(Vector[num].name2);
	if (num < k) {
		this->infoUcoin = allInformation.substr(positionBegin, allInformation.find(Vector[num + 1].name2) - positionBegin);
	}
	else {
		this->infoUcoin = allInformation.substr(positionBegin, allInformation.length() - positionBegin);
	}
};





void money::GetInfoFromThreeSite()
{
	system("cls");
	cout << "Номер монеты: " << number << " " << name2 << endl;
	cout << "Поиск информации ucoin.ru: ";
	try {
		GetWeightDiameterThicknessUcoin();
		GetPriceEdition();
		GetPriceUcoin();
		GetMaterialUcoin();
		cout << endl << "material: " << material;
		cout << "coin_edge: " << coin_edge << endl << endl;
		cout << "OK!" << endl;
	}
	catch (...) { cout << "Нет информации" << endl; }
	cout << "Поиск информации raritetus.ru: ";
	try {
		this->htmlRaritetus = GetDataFromSite("https://www.raritetus.ru/search/catalog/?par=" + this->url);
		ParseUrlMoneyRaritetus();
		this->htmlRaritetus = GetDataFromSite("https://www.raritetus.ru" + this->urlRaritetus);
	}
	catch (...) {}
	if (this->htmlRaritetus != "" && this->urlRaritetus != "") {
		try {
			this->GetInfoFromRaritetus();
			cout << "OK!" << endl;
		}
		catch (...) { cout << "Нет информации" << endl; }

	}
	cout << "Поиск информации coinsmart.ru: ";
	try {
		this->htmlCoinsmart = GetDataFromSite("https://coinsmart.ru/search/?query=" + this->url);
		ParseUrlMoneyCoinsmart();
		this->htmlCoinsmart = GetDataFromSite("https://coinsmart.ru" + this->urlCoinsmart);
	}
	catch (...) {}
	if (this->htmlCoinsmart != "" && this->urlCoinsmart != "") {
		try {
			this->GetInfoFromCoinsmart();
			cout << "OK!" << endl;
		}
		catch (...) { cout << "Нет информации" << endl; }
	}

}


//--------------------------------------------------//

int main() {
	setlocale(LC_ALL, "");
	SetConsoleCP(1251);
	SetConsoleOutputCP(1251);

	int k = count_money();
	cout << "Количество монет в файле: " << k << endl;
	string allInformation;
	string allMoney;
	vector <money> moneyVector;
	for (int i = 1; i <= k; i++) {
		money M(i);
		M.GetMoneyFromExcel();
		M.name2 = UTF8_to_CP1251(M.name);
		moneyVector.push_back(M);
		if (i != k)
			allMoney += M.name += '\n';
		else
			allMoney += M.name;
	}

	bool parseUcoin;
	cout << "Обновить файл информации с Ucoin? (1 - да, 0 - нет): ";
	cin >> parseUcoin;

	if (parseUcoin)
	{
		ofstream out;
		out.open("money.txt");
		out << allMoney;
		out.close();
		system("java -jar OldMoneyParser.jar");
	}

	ifstream in;
	in.open("all information.txt");
	std::stringstream ss;
	ss << in.rdbuf();
	allInformation = ss.str();
	in.close();

	for (int i = 0; i < k; i++) {
		try {
			moneyVector[i].CutStringAllMoney(moneyVector, i, k, allInformation);
			moneyVector[i].CutStringFromUcoin();
		}
		catch (...) { cout << "Error with cut!"; }
		try {
			moneyVector[i].GetInfoFromThreeSite();
			moneyVector[i].PostMoneyIntoExcel();
		}
		catch (...) {}
	}


	return 0;
}