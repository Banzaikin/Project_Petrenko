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
	doc.open("./money.xlsx");
	auto wks = doc.workbook().worksheet("Sheet1");
	int B1 = wks.cell("B1").value();
	doc.close();
	return B1;
}

//получение информации о монете по порядковому номеру
string* get_money(int number)
{
	string num = to_string(number+2);
	XLDocument doc;
	doc.open("./money.xlsx");
	auto wks = doc.workbook().worksheet("Sheet1");
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
void post_money(int price, float weight, float diametr, int number)
{
	string num = to_string(number + 2);
	XLDocument doc;
	doc.open("./money.xlsx");
	auto wks = doc.workbook().worksheet("Sheet1");
	wks.cell("R" + num).value() = price;
	wks.cell("J" + num).value() = weight;
	wks.cell("L" + num).value() = diametr;
	doc.save();
	cout << num << ":   " << "Average cost: " << price << "   Weight: " << weight << "   Diameter: " << diametr << endl;
	doc.close();
}

//получение адреса страницы монеты из страницы поиска
string parse_url_money(string html)
{
	int StUrl, EnUrl;
	string NewUrl;
	StUrl = html.find("url='/stoimost-monet") + 5;
	EnUrl = html.find("' class=");
	NewUrl = html.substr(StUrl, EnUrl-StUrl);
	return NewUrl;
}

//извлечение чисел в вектор из строки
vector<int> num_from_string(string str)
{
	vector <int> ivector;
	char temp[1024];
	strcpy(temp, str.c_str());
	for (auto i = strtok(temp, " \f\n\r\t\v"); i != nullptr; i = strtok(nullptr, " \f\n\r\t\v")) {
		char* it;
		double num = strtod(i, &it);

		if (*it == '\0') {
			ivector.push_back(num);
		}
	}
	return ivector;
}

//получение средней цены по коду страницы и сохранности
int get_middle_price(string html, string condition_str) {
	std::ofstream out;          
	out.open("hello2.txt");

	int middlePrice, StPrices, numbCond;
	string condidion_mas[] = { "G","VG","F","VF","XF","AU","UNC" };
	for (int i = 0; i < 7; i++)
		if (condidion_mas[i] == condition_str) {
			numbCond = i;
			break;
		}
	StPrices = html.find("avg-prices");
	string Prices = html.substr(StPrices+106, 94);

	while (Prices.find("-") != std::string::npos) {
		Prices.replace(Prices.find("-"), 1, "0");
	}
	while (Prices.find(" ") != std::string::npos) {
		Prices.erase(Prices.find(" "), 1);
	}
	vector<int> condition_vec = num_from_string(Prices);
	out.close();
	return condition_vec[numbCond];
}

//получение веса монеты
float get_weight_money(string html)
{
	std::ofstream out;
	out.open("hello2.txt");

	int weightMoneyStr = html.find("col-sm-4 descfullcont");
	string weight = html.substr(weightMoneyStr+230, 50);
	char tab2[50], tabWeight[4];
	strcpy(tab2, weight.c_str());
	for (int i = 0; i < 50; i++)
	{
		if (isdigit(tab2[i]))
		{
			tabWeight[0] = tab2[i];
			i++;
			if ((tab2[i]) == '.' || (tab2[i]) == ',')
			{
				for (int j = 1; j < 4; j++)
				{
					if (isdigit(tab2[i]) || tab2[i] == '.' || tab2[i] == ',')
					{
						if (tab2[i] == ',')
							tab2[i] = '.';
						tabWeight[j] = tab2[i];
					}
					else
						tabWeight[j] = '0';
					i++;
				}
			}
		}
	}
	auto weightMoney = atof(tabWeight);
	out.close();
	return weightMoney;
}

//получение диаметра монеты
float get_diametr_money(string html)
{
	std::ofstream out;
	out.open("hello2.txt");

	float diametrMoney;
	int diametrMoneyStr = html.find("col-sm-4 descfullcont");
	string diametr = html.substr(diametrMoneyStr + 270, 50);
	char tab2[50], tabDiametr[4];
	strcpy(tab2, diametr.c_str());
	for (int i = 0; i < 50; i++)
	{
		if (isdigit(tab2[i]))
		{
			tabDiametr[0] = tab2[i];
			i++;
			tabDiametr[1] = tab2[i];
			i++;	
			if (tab2[i] == '.' || tab2[i] == ',')
			{
				if (tab2[i] == ',')
					tab2[i] = '.';
				tabDiametr[2] = tab2[i];
				i++;
				tabDiametr[3] = tab2[i];
				i++;
			}
		}
	}
	diametrMoney = atof(tabDiametr);
	out.close();
	return diametrMoney;
}

int main() {
	int k = count_money();
	cout << "Number of coins: " << k << endl;
	for (int i = 1; i <= k; i++) {
		string* m;
		string money, condition, html, url;
		m = get_money(i);
		money = m[0];
		condition = m[1];
		html = get_data_from_site("https://www.raritetus.ru/search/catalog/?par=" + money);
		if (html == "")
			break;
		url = parse_url_money(html);
		html = get_data_from_site("https://www.raritetus.ru" + url);
		if (html == "")
			break;
		int price =  get_middle_price(html, condition); 
		float weight = get_weight_money(html);
		float diametr = get_diametr_money(html);
		post_money(price, weight, diametr, i);
	}
	return 0;
}