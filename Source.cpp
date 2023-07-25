#define CURL_STATICLIB
#define _CRT_SECURE_NO_WARNINGS
#pragma once
#include <vector>
#include <string>
#include <iostream>
#include "curl\curl.h"
#include <windows.h>
#include <fstream>
#include <cmath>
#include <OpenXLSX\OpenXLSX.hpp>
#include <stdlib.h>

using namespace OpenXLSX;
using namespace std;

class money {
	int number;
	string condition;
	double weight = 0;
	double diametr = 0;
	int edition = 0;
	int middlePrice1 = 0;
	int middlePrice2 = 0;
public:
	string url;
	string url2;
	string html1;
	string html2;
	string html3;
	money(int num) { number = num; }
	void get_money();
	void post_money();
	void get_middle_price();
	void get_weight_diameter_edition_money();
	void price_money2();
	void get_weight_diameter_edition_money2();
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

	std::ofstream out;          // поток для записи
	out.open("hello.txt");      // открываем файл для записи
	out << url;
	out << data;
	out.close();

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

	url = C + " " + D2 + " " + F;

	while (url.find(" ") != string::npos) {
		url.replace(url.find(" "), 1, "+");
	}
	string E = wks.cell("E" + num).value();
	condition = E;
	doc.close();
}

//запись даннных о монете
void money::post_money()
{
	string num = to_string(number + 2);
	XLDocument doc;
	doc.open("./111.xlsx");
	auto wks = doc.workbook().worksheet("Main");
	wks.cell("R" + num).value() = middlePrice1;
	wks.cell("S" + num).value() = middlePrice2;
	wks.cell("J" + num).value() = weight;
	wks.cell("L" + num).value() = diametr;
	wks.cell("K" + num).value() = edition;

	doc.save();
	cout << num << ":   " << "1 price: " << middlePrice1 << " 2 price: " << middlePrice2 << "   Weight: " << weight << "   Edition: " << edition << "   Diameter: " << diametr << endl;
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
//--------------------------------------------------//

int main() {
	int k = count_money();
	cout << "Number of coins: " << k << endl;
	for (int i = 1; i <= k; i++) {
		cout << i << " ";
		money M(i);
		M.get_money();
		M.html1 = get_data_from_site("https://www.raritetus.ru/search/catalog/?par=" + M.url);
		M.url2 = parse_url_money(M.html1);
		M.html1 = get_data_from_site("https://www.raritetus.ru" + M.url2);
		if (M.html1 != "" && M.url2 != "") {
			M.get_middle_price();
			M.get_weight_diameter_edition_money();
			M.post_money();
		}
		else {
			cout << "Not info1";
		}
		cout << '\n';
		cout << "Look at the site coinsmart.ru ..." << endl;
		M.html2 = get_data_from_site("https://coinsmart.ru/search/?query=" + M.url);
		M.url2 = parse_url_money_2(M.html2);
		M.html2 = get_data_from_site("https://coinsmart.ru" + M.url2);
		if (M.html2 != "" && M.url2 != "") {
			M.price_money2();
			M.get_weight_diameter_edition_money2();
			M.post_money();
		}
		else {
			cout << "Not info2";
		}
		cout << '\n' << "ucoin";
		M.html3 = get_data_from_site("https://ru.ucoin.net/catalog/?q=" + M.url);
		M.url2 = parse_url_money_3(M.html3);
		M.html3 = get_data_from_site("https://ru.ucoin.net" + M.url2);
		cout << "end\n";
		Sleep(5000);
	}
	return 0;
}
