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
	double year;
	int number;
	string condition;
	double weight = 0;
	double diametr = 0;
	int edition = 0;
	int middlePrice1 = 0;
	int middlePrice2 = 0;
	double middlePrice3 = 0;
public:
	string url;
	string url2;
	string html1;
	string html2;
	string html3;
	money(int num) { number = num; };
	money() {};
	void get_money();
	void post_money();
	void get_middle_price();
	void get_weight_diameter_edition_money();
	void price_money2();
	void get_weight_diameter_edition_money2();
	void get_weight_diameter_edition_money3();
	void GetWeightDiameterThicknessUcoin();
	void CutStringAllMoney(vector <money> Vector, int num, int k, string allInformation);
	void GetInfoFromThreeSite();
	void CutStringFromUcoin();
};
size_t write_data(void* ptr, size_t size, size_t nmemb, std::string* data) {
	data->append((char*)ptr, size * nmemb);
	return size * nmemb;
}

// получение html по ссылке
string get_data_from_site(string url) {
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
void money::get_money()
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
void money::post_money()
{
	string num = to_string(number + 2);
	XLDocument doc;
	doc.open("./111.xlsx");
	auto wks = doc.workbook().worksheet("Main");
	wks.cell("T" + num).value() = middlePrice1;
	wks.cell("U" + num).value() = middlePrice2;
	wks.cell("V" + num).value() = middlePrice3;
	wks.cell("K" + num).value() = weight;
	wks.cell("N" + num).value() = diametr;
	wks.cell("L" + num).value() = edition;
	wks.cell("M" + num).value() = thickness;

	doc.save();
	//cout << num << ":   " << "1 price: " << middlePrice1 << " 2 price: " << middlePrice2 << "   Weight: " << weight << "   Edition: " << edition << "   Diameter: " << diametr << endl;
	doc.close();
}

//получение адреса страницы монеты из страницы поиска
string parse_url_money(string html)
{
	int StUrl, EnUrl;
	string NewUrl;
	StUrl = html.find("url='/stoimost-monet") + 5;
	EnUrl = html.find("' class=");
	if (StUrl == 4 && EnUrl == -1)
		return "";
	else {
		NewUrl = html.substr(StUrl, EnUrl - StUrl);
		return NewUrl;
	}
}

template <typename T>
//извлечение чисел в вектор из строки
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

//получение средней цены по коду страницы и сохранности
void money::get_middle_price() {
	int StPrices, numbCond;
	string condidion_mas[] = { "G","VG","F","VF","XF","AU","UNC" };
	for (int i = 0; i < 7; i++)
		if (condidion_mas[i] == condition) {
			numbCond = i;
			break;
		}
	StPrices = html1.find("avg-prices");
	string Prices = html1.substr(StPrices + 106, 94);

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
void money::get_weight_diameter_edition_money()
{
	int MoneyStr = html1.find("col-sm-4 descfullcont");

	string edition_s = html1.substr(MoneyStr - 80, 35);
	while (edition_s.find(" ") != std::string::npos) {
		edition_s.erase(edition_s.find(" "), 1);
	}
	int i = 1;
	vector<int> vec1 = num_from_string(edition_s, i);
	if (vec1.empty())
		edition = 0;
	else
		edition = vec1[0];

	string weight_s = html1.substr(MoneyStr + 230, 150);
	while (weight_s.find(",") != std::string::npos) {
		weight_s.replace(weight_s.find(","), 1, ".");
	}
	double d = 0.1;
	vector<double> vec2 = num_from_string(weight_s, d);

	weight = vec2[0];
	diametr = vec2[1];
}


//-----------------------------------------//
//Функции для 2 сайта

//получение адреса страницы монеты из страницы поиска
string parse_url_money_2(string html)
{
	int StUrl, EnUrl;
	string NewUrl;
	StUrl = html.find("data-url=") + 10;
	EnUrl = html.find("?cart=");
	NewUrl = html.substr(StUrl, EnUrl - StUrl);
	return NewUrl;
}

//получение цены с аукциона
void money::price_money2()
{
	int priceMoneyStr = html2.find("data-price=");
	string priceMoney = html2.substr(priceMoneyStr, 50);

	while (priceMoney.find(" ") != std::string::npos) {
		priceMoney.erase(priceMoney.find(" "), 1);
	}

	int d = 1;
	vector<int> vec = num_from_string(priceMoney, d);

	middlePrice2 = vec[0];
}

//получение диаметра, веса и тиража
void money::get_weight_diameter_edition_money2()
{
	int weightMoneyStr = html2.find("features striped");
	string weight_s = html2.substr(weightMoneyStr + 250, 400);

	while (weight_s.find(",") != std::string::npos) {
		weight_s.replace(weight_s.find(","), 1, ".");
	}

	while (weight_s.find(" ") != std::string::npos) {
		weight_s.erase(weight_s.find(" "), 1);
	}

	double d = 0.1;
	vector<double> vec = num_from_string(weight_s, d);

	if (weight == 0)
		weight = vec[1];
	if (edition == 0)
		edition = vec[3];
	if (diametr == 0)
		diametr = vec[0];

}

string parse_url_money_3(string html)
{
	int StUrl, EnUrl;
	string SubHtml, NewUrl;
	StUrl = html.find("</script></center>");
	SubHtml = html.substr(StUrl, 500);
	StUrl = SubHtml.find("/coin/");
	while (SubHtml[StUrl] != '"')
	{
		NewUrl += SubHtml[StUrl];
		StUrl++;
	}
	cout << NewUrl;
	return NewUrl;
}


//void money::get_weight_diameter_edition_money3()
//{
//
//	// таблица с разновидностью
//	try {
//		int positionTable1 = text.find("Знак Описание");
//		if (positionTable1 != string::npos)
//		{
//			cout << "!!!!!!!";
//			string tableText = text.substr(positionTable1, text.length() - positionTable1);
//			int positionToken = tableText.find(this->token);
//			cout << positionToken;
//			if (positionToken != string::npos)
//			{
//				string rowsWithToken = tableText.substr(positionToken, tableText.length() - positionToken);
//				string rowWithToken = rowsWithToken.substr(0, rowsWithToken.find("\n"));
//				double i = 1.0;
//				vector<double> vec1 = num_from_string(rowWithToken, i);
//				this->middlePrice3 = vec1[0];
//				cout << rowWithToken;
//			}
//
//		}
//	}
//	catch (...) { this->middlePrice3 = 0; }
//
//
//
//	// таблица с тиражом
//	int positionTable2 = text.find("Тираж");
//	for (int i = positionTable2; i < text.length(); i++) {
//		if (text[i] == '\n' && text[i + 1] == '\n')
//		{
//			text[i] = '\n';
//			text[i + 1] = 'V';
//		}
//	}
//	//cout << text;
//	string textWithYear = text.substr(positionTable2, text.find('V') - positionTable2);
//	int positionYear = textWithYear.find(to_string(year));
//	if (positionYear == string::npos)
//	{
//		try
//		{
//			if (text.find("Цена") != string::npos) {
//				double j = 1;
//				vector<double> vec2 = num_from_string(textWithYear, j);
//				this->middlePrice3 = vec2[vec2.size() - 1];
//			}
//			int pos = textWithYear.find(".");
//			while (pos != string::npos)
//			{
//				textWithYear.replace(pos, 1, "");
//				pos = textWithYear.find(".");
//			}
//			int i = 1;
//			vector<int> vec1 = num_from_string(textWithYear, i);
//			if (vec1[0] == year) {
//				this->edition = vec1[1];
//			}
//			else{ this->edition = vec1[0]; }
//		}
//		catch (...){}
//	}
//	else
//	{
//
//		string rowsWithYear = textWithYear.substr(positionYear, textWithYear.length());
//		string rowWithYear = rowsWithYear.substr(0, rowsWithYear.find("\n"));
//		try
//		{
//			double j = 1;
//			vector<double> vec2 = num_from_string(rowWithYear, j);
//			if ((vec2[vec2.size() - 1]) != year)
//			{
//				this->middlePrice3 = vec2[vec2.size() - 1];
//
//				int pos = rowWithYear.find(".");
//				while (pos != string::npos)
//				{
//					rowWithYear.replace(pos, 1, "");
//					pos = rowWithYear.find(".");
//				}
//
//				int i = 1;
//				vector<int> vec1 = num_from_string(rowWithYear, i);
//				this->edition = vec1[1];
//			}
//		}
//		catch (...){}
//	}
//
//
//}

void money::GetWeightDiameterThicknessUcoin() {
	string text = vectorCutStringUcoin[1];
	int position1 = text.find("Вес");
	string numText = text.substr(position1, text.length() - position1);
	try {
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
		infoUcoinText.erase(0, infoUcoinText.find('№')+1);
	}
}

void money::CutStringAllMoney(vector <money> Vector, int num, int k, string allInformation) {
	int positionBegin = allInformation.find(Vector[num].name2);
	if (num < k) {
		this->infoUcoin = allInformation.substr(positionBegin, allInformation.find(Vector[num+1].name2) - positionBegin);
	}
	else {
		this->infoUcoin = allInformation.substr(positionBegin, allInformation.length() - positionBegin);
	}
};

void money::GetInfoFromThreeSite()
{
	system("cls");
	cout << "Money number: " << number << " " << name2 << endl;
	/*cout << "Find info from raritetus.ru: ";
	try {
		this->html1 = get_data_from_site("https://www.raritetus.ru/search/catalog/?par=" + this->url);
		this->url2 = parse_url_money(this->html1);
		this->html1 = get_data_from_site("https://www.raritetus.ru" + this->url2);
	}
	catch ( ... ) {}
	if (this->html1 != "" && this->url2 != "") {
		try {
			this->get_middle_price();
			this->get_weight_diameter_edition_money();
			cout << "OK!" << endl;
		}
		catch (...) { cout << "Not info" << endl; }

	}
	cout << "Find info from coinsmart.ru: " << endl;*/
	/*try {
		this->html2 = get_data_from_site("https://coinsmart.ru/search/?query=" + this->url);
		this->url2 = parse_url_money_2(this->html2);
		this->html2 = get_data_from_site("https://coinsmart.ru" + this->url2);
	}
	catch (...) {}
	if (this->html2 != "" && this->url2 != "") {
		try {
			this->price_money2();
			this->get_weight_diameter_edition_money2();
			cout << "OK!" << endl;
		}
		catch (...) { cout << "Not info" << endl; }
	}*/
	cout << "Find info from ucoin.ru: " << endl;
	try {
		GetWeightDiameterThicknessUcoin();
		cout << "OK!" << endl;
	}
	catch (...) { cout << "Not info" << endl; }
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
//--------------------------------------------------//

int main() {
	SetConsoleCP(1251); 
	SetConsoleOutputCP(1251);
	int k = count_money();
	cout << "Number of coins: " << k << endl;
	string allInformation;
	string allMoney;
	vector <money> moneyVector;
	for (int i = 1; i <= k; i++) {
		money M(i);
		M.get_money();
		M.name2 = UTF8_to_CP1251(M.name);
		moneyVector.push_back(M);
		if (i!=k)
			allMoney += M.name += '\n';
		else
			allMoney += M.name;
	}

	/*ofstream out;         
	out.open("money.txt");      
	out << allMoney;
	out.close();
	system("java -jar OldMoneyParser.jar");*/

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
			moneyVector[i].post_money();
		}
		catch( ... ) {}
	}

	return 0;
}
