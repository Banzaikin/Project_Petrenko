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

size_t write_data(void* ptr, size_t size, size_t nmemb, std::string* data) {
	data->append((char*)ptr, size * nmemb);
	return size * nmemb;
}

// получение html по ссылке
std::string get_data_from_site(std::string url) {
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
	auto wks = doc.workbook().worksheet("Основное");
	int B1 = wks.cell("B1").value();
	doc.close();
	return B1;
}

//получение информации о монете по порядковому номеру
string* get_money(int number)
{
	string num = to_string(number + 2);
	XLDocument doc;
	doc.open("./111.xlsx");
	auto wks = doc.workbook().worksheet("Основное");
	string C = wks.cell("C" + num).value();
	int D = wks.cell("D" + num).value();
	string D2 = to_string(D);
	string F = wks.cell("F" + num).value();

	C = C + D2 + F;

	while (C.find(" ") != std::string::npos) {
		C.erase(C.find(" "), 1);
	}
	string E = wks.cell("E" + num).value();

	doc.close();

	string* res = new string[2];
	res[0] = C;
	res[1] = E;
	return res;
}

//запись даннных о монете
void post_money(int price, float weight, int edition, float diametr, int number)
{
	string num = to_string(number + 2);
	XLDocument doc;
	doc.open("./111.xlsx");
	auto wks = doc.workbook().worksheet("Основное");
	wks.cell("R" + num).value() = price;
	wks.cell("J" + num).value() = weight;
	wks.cell("L" + num).value() = diametr;
	wks.cell("K" + num).value() = edition;
	doc.save();
	cout << num << ":   " << "Average cost: " << price << "   Weight: " << weight << "   Edition: " << edition << "   Diameter: " << diametr << endl;
	doc.close();
}

//получение адреса страницы монеты из страницы поиска
string parse_url_money(string html)
{
	int StUrl, EnUrl;
	string NewUrl;
	StUrl = html.find("url='/stoimost-monet") + 5;
	EnUrl = html.find("' class=");
	NewUrl = html.substr(StUrl, EnUrl - StUrl);
	return NewUrl;
}

template <typename T>
//извлечение чисел в вектор из строки
vector<T> num_from_string(string str, T x)
{
	vector <T> Tvector;
	char temp[1024];
	strcpy(temp, str.c_str());
	for (auto i = strtok(temp, " \f\n\r\t\v<>"); i != nullptr; i = strtok(nullptr, " \f\n\r\t\v<>")) {
		char* it;
		double num = strtod(i, &it);

		if (*it == '\0') {
			Tvector.push_back(num);
		}
	}
	return Tvector;
}

//получение средней цены по коду страницы и сохранности
int get_middle_price(string html, string condition_str) {
	std::ofstream out;
	out.open("hello2.txt");

	int StPrices, numbCond;
	string condidion_mas[] = { "G","VG","F","VF","XF","AU","UNC" };
	for (int i = 0; i < 7; i++)
		if (condidion_mas[i] == condition_str) {
			numbCond = i;
			break;
		}
	StPrices = html.find("avg-prices");
	string Prices = html.substr(StPrices + 106, 94);

	while (Prices.find("-") != std::string::npos) {
		Prices.replace(Prices.find("-"), 1, "0");
	}
	while (Prices.find(" ") != std::string::npos) {
		Prices.erase(Prices.find(" "), 1);
	}
	vector<int> condition_vec = num_from_string(Prices, numbCond);
	out.close();
	return condition_vec[numbCond];
}

//получение веса и диаметра монеты
vector<double> get_weight_diameter_money(string html)
{
	int weightMoneyStr = html.find("col-sm-4 descfullcont");
	string weight = html.substr(weightMoneyStr + 230, 150);

	while (weight.find(",") != std::string::npos) {
		weight.replace(weight.find(","), 1, ".");
	}

	double d = 0.1;
	vector<double> vec = num_from_string(weight, d);
	return vec;
}

//получение тиража монеты
int get_edition_money(string html) {

	int editionMoneyStr = html.find("col-sm-4 descfullcont");
	string edition = html.substr(editionMoneyStr - 80, 35);

	while (edition.find(" ") != std::string::npos) {
		edition.erase(edition.find(" "), 1);
	}

	int d = 1;
	vector<int> vec = num_from_string(edition, d);
	if (vec.empty())
		return 0;
	else
		return vec[0];
}

//-----------------------------------------//
//Функции для 2 сайта

//получение информации о монете по порядковому номеру
string get_money_2(int number)
{
	string num = to_string(number + 2);
	XLDocument doc;
	doc.open("./111.xlsx");
	auto wks = doc.workbook().worksheet("Основное");
	string C = wks.cell("C" + num).value();
	int D = wks.cell("D" + num).value();
	string D2 = to_string(D);
	string F = wks.cell("F" + num).value();

	C = C + D2 + F;

	while (C.find(" ") != std::string::npos) {
		C.replace(C.find(" "), 1, "+");
	}

	doc.close();

	return C;
}

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
int price_money(string html)
{
	int priceMoneyStr = html.find("data-price=");
	string priceMoney = html.substr(priceMoneyStr, 50);

	while (priceMoney.find(" ") != std::string::npos) {
		priceMoney.erase(priceMoney.find(" "), 1);
	}

	int d = 1;
	vector<int> vec = num_from_string(priceMoney, d);
	if (vec.empty())
		return 0;
	else
		return vec[0];
}

//получение диаметра, веса и тиража
vector<double> get_weight_diameter_edition_money(string html)
{
	int weightMoneyStr = html.find("features striped");
	string weight = html.substr(weightMoneyStr + 250, 400);

	while (weight.find(",") != std::string::npos) {
		weight.replace(weight.find(","), 1, ".");
	}

	while (weight.find(" ") != std::string::npos) {
		weight.erase(weight.find(" "), 1);
	}

	double d = 0.1;
	vector<double> vec = num_from_string(weight, d);
	auto iter = vec.cbegin();
	if (vec.size() < 4)
		throw "Error: Not found on the site";
	return vec;
}

//запись даннных о монете
void post_money_2(int price, float weight, int edition, float diametr, int number)
{
	string num = to_string(number + 2);
	XLDocument doc;
	doc.open("./111.xlsx");
	auto wks = doc.workbook().worksheet("Основное");
	wks.cell("S" + num).value() = price;
	wks.cell("J" + num).value() = weight;
	wks.cell("L" + num).value() = diametr;
	wks.cell("K" + num).value() = edition;
	doc.save();
	cout << num << ":   " << "Average cost: " << price << "   Weight: " << weight << "   Edition: " << edition << "   Diameter: " << diametr << endl;
	doc.close();
}

//--------------------------------------------------//

int main() {
	int k = count_money();
	cout << "Number of coins: " << k << endl;
	for (int i = 1; i <= k; i++) {
		string* m;
		string money, condition, html, url;
		//для 1 сайта
		m = get_money(i);
		money = m[0];
		condition = m[1];
		html = get_data_from_site("https://www.raritetus.ru/search/catalog/?par=" + money);
		url = parse_url_money(html);
		html = get_data_from_site("https://www.raritetus.ru" + url);
		if (html != "")
		{
			int price = get_middle_price(html, condition);
			vector<double> weightAndDiameter = get_weight_diameter_money(html);
			int edition = get_edition_money(html);
			post_money(price, weightAndDiameter[0], edition, weightAndDiameter[1], i);
		}
		else
		{
			//для 2 сайта
			cout << "Look at the site coinsmart.ru ..." << endl;
			money = get_money_2(i);
			html = get_data_from_site("https://coinsmart.ru/search/?query=" + money);
			url = parse_url_money_2(html);
			html = get_data_from_site("https://coinsmart.ru" + url);
			cout << endl << url << endl << endl;
			try {
				if (html == "")
				{
					cout << "Not info" << endl;
					continue;
				}
				int price = price_money(html);
				vector<double> weightDiameterEdition = get_weight_diameter_edition_money(html);
				post_money_2(price, weightDiameterEdition[1], weightDiameterEdition[3], weightDiameterEdition[0], i);
			}
			catch (const char* error_message){
				cout << error_message << endl;
			}
		}
	}
	return 0;
}
