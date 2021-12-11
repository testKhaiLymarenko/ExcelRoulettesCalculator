package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"strconv"
	"strings"
	"time"

	"github.com/fatih/color"
	"github.com/xuri/excelize/v2"
)

func writeToFileMonthIncome(exFile *excelize.File, accounts *[]Accounts, totalWtfskinsIncome, totalCsgolivesIncome,
	totalPvproDollarsIncome, totalOverallIncomeInDollars float64, totalPvproCoinsIncome int, fileMonthName string) {
	incomeSheetName := exFile.GetSheetName(0)

	for i := 'B'; i <= 'D'; i++ {
		for j := range *accounts {

			if i == 'B' {
				err := exFile.SetCellValue(incomeSheetName, "B"+strconv.Itoa(j+2),
					fmt.Sprintf("+$%.2f", (*accounts)[j].wtfskinsIncome))

				if err != nil {
					fmt.Println(err)
				}
			} else if i == 'C' {
				err := exFile.SetCellValue(incomeSheetName, "C"+strconv.Itoa(j+2),
					fmt.Sprintf("+$%.2f", (*accounts)[j].csgoliveIncome))

				if err != nil {
					fmt.Println(err)
				}
			} else if i == 'D' {
				err := exFile.SetCellValue(incomeSheetName, "D"+strconv.Itoa(j+2),
					"+"+strconv.Itoa((*accounts)[j].pvproCoinsIncome)+" coins")

				if err != nil {
					fmt.Println(err)
				}
			}
		}
	}

	if err := exFile.Save(); err != nil {
		fmt.Println(err)
	}

	if err := exFile.SetCellValue(incomeSheetName, "B8", fmt.Sprintf("+$%.2f", totalWtfskinsIncome)); err != nil {
		fmt.Println(err)
	}

	if err := exFile.SetCellValue(incomeSheetName, "C8", fmt.Sprintf("+$%.2f", totalCsgolivesIncome)); err != nil {
		fmt.Println(err)
	}

	if err := exFile.SetCellValue(incomeSheetName, "D8", fmt.Sprintf("+%d coins (+$%.2f)", totalPvproCoinsIncome, totalPvproDollarsIncome)); err != nil {
		fmt.Println(err)
	}

	if err := exFile.SetCellValue(incomeSheetName, "C11", fmt.Sprintf("Total Income: $%.2f", totalOverallIncomeInDollars)); err != nil {
		fmt.Println(err)
	}

	if err := exFile.Save(); err != nil {
		fmt.Println(err)
	} else {
		color.Green("\n\nCalculated values have been successfully stored")
	}
}

func writeToFileOverallIncome(workFolder, fileTotalIncomeName, fileMonthName string, totalOverallIncomeInDollars float64,
	incomeMonth, incomeYear int) {

	exFileIncome, err := excelize.OpenFile(workFolder + "\\" + fileTotalIncomeName)

	if err != nil {
		fmt.Println(err)
		fmt.Scanln()
		return
	}

	sheetsNumber := 0
	for {
		incomeSheetName := exFileIncome.GetSheetName(sheetsNumber)

		if incomeSheetName == "" {
			break
		}

		sheetsNumber++
	}

	//2020 - first sheet (at index 0)
	sheetYearIndex := incomeYear - 2020

	var date string

	//if we check roulettes data of previous month we choose 28th day, if current month - so current day
	if time.Now().Month() == time.Month(incomeMonth) && time.Now().Day() <= 28 {
		date = time.Now().Format("02.01.2006")
	} else {
		//28.11.2021, 28.02.2022, 28.05.2022 -> cuz 28th day exists in each month
		date = "28." + fmt.Sprintf("%02d", incomeMonth) + "." + strconv.Itoa(incomeYear)
	}

	resp, err := http.Get("https://api.privatbank.ua/p24api/exchange_rates?json&date=" + date)

	if err != nil {
		fmt.Println(err)
		fmt.Scanln()
		return
	}

	defer resp.Body.Close()

	data, _ := ioutil.ReadAll(resp.Body)

	//remove useless from the beginning ""date":"09.12.2021","bank":"PB","baseCurrency":980,"baseCurrencyLit":"UAH","exchangeRate":["
	//privatData := []byte(string(buff)[strings.Index(string(buff), "[")+1:])
	data = []byte(string(data)[strings.Index(string(data), "[") : len(data)-1])

	//json line: "baseCurrency":"UAH","currency":"RUB","saleRateNB":0.3692500,"purchaseRateNB":0.3692500,"saleRate":0.3860000,"purchaseRate":0.3560000
	type Currency struct {
		ForeignCurrencyName string  `json:"currency"`
		PurchaseRateNBU     float64 `json:"purchaseRateNB"`
	}

	currencies := []Currency{}
	if err = json.Unmarshal(data, &currencies); err != nil {
		fmt.Println(err)
		fmt.Scanln()
		return
	}

	var totalOverallIncomeInHryvnia, totalOverallIncomeInRubles float64

	for _, currency := range currencies {
		if currency.ForeignCurrencyName == "USD" {
			totalOverallIncomeInHryvnia = currency.PurchaseRateNBU * totalOverallIncomeInDollars
		}
	}

	//2 loops cuz we need to get income in hryvnia first
	for _, currency := range currencies {
		if currency.ForeignCurrencyName == "RUB" {
			totalOverallIncomeInRubles = totalOverallIncomeInHryvnia / currency.PurchaseRateNBU
		}
	}

	//Set value for the Dollars cell
	if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "B"+strconv.Itoa(incomeMonth+1),
		fmt.Sprintf("$%.2f", totalOverallIncomeInDollars)); err != nil {

		fmt.Println(err)
		fmt.Scanln()
		return

	}

	//Set value for the Rubles cell
	if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "C"+strconv.Itoa(incomeMonth+1),
		fmt.Sprintf("₽%.2f", totalOverallIncomeInRubles)); err != nil {

		fmt.Println(err)
		fmt.Scanln()
		return

	}

	//Set value for the Hryvnia cell
	if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "D"+strconv.Itoa(incomeMonth+1),
		fmt.Sprintf("₴%.2f", totalOverallIncomeInHryvnia)); err != nil {

		fmt.Println(err)
		fmt.Scanln()
		return

	}

	if err := exFileIncome.Save(); err != nil {
		fmt.Println(err)
	} else {
		color.Green("\n\nCalculated values have been successfully stored")
	}
}
