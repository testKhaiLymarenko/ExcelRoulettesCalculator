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

//брать курс валют с НБУ
func main() {

	workFolder := "D:\\Program Files\\MEGAsync\\MEGAsync\\Internet Deals\\Steam\\ルーレット"
	fileTotalIncomeName := "_ルーレットの総収入.xlsx"

	/*fmt.Print(workFolder + ": ")

	fileMonthName := bufio.NewScanner(os.Stdin) //2021年12月のルーレット.xlsx
	fileMonthName.Scan()

	if fileMonthName.Err() != nil {
		fmt.Println(fileMonthName.Err())
		fmt.Scanln()
		return
	}*/

	//exFile, err := excelize.OpenFile(workFolder + "\\" + fileMonthName.Text())
	exFile, err := excelize.OpenFile(workFolder + "\\" + "2021年12月のルーレット.xlsx")

	if err != nil {
		fmt.Println(err)
		fmt.Scanln()
		return
	}

	accounts := []Accounts{
		{login: "l......_"},
		{login: ".....l"},
		{login: "de....9"},
		{login: "d...."},
		{login: "....."},
	}

	var totalWtfskinsIncome, totalCsgolivesIncome, totalPvproDollarsIncome float64
	var totalPvproCoinsIncome int
	var totalOverallIncomeInDollars float64

	//both functions store data to the struct
	checkAndGetFirstCells(exFile, &accounts)
	getLastCellValues(exFile, &accounts)

	//Print income of each account and count the total income in loop to print it later
	for i := range accounts {
		switch accounts[i].login { //index is needed cuz range-loop copies accounts[i] to account, but not a pointer
		case "....._":
			color.Red(accounts[i].CalculateS())
		case "r.....":
			color.Magenta(accounts[i].CalculateS())
		case "d...":
			color.Yellow(accounts[i].CalculateS())
		case "de....":
			color.Cyan(accounts[i].CalculateS())
		case "d......2":
			color.White(accounts[i].CalculateS())
		}

		totalWtfskinsIncome += accounts[i].wtfskinsIncome
		totalCsgolivesIncome += accounts[i].csgoliveIncome

		totalPvproCoinsIncome += accounts[i].pvproCoinsIncome
		totalPvproDollarsIncome += accounts[i].pvproDollarsIncome

	}

	//print Total Income
	totalOverallIncomeInDollars = totalWtfskinsIncome + totalCsgolivesIncome + totalPvproDollarsIncome
	color.Green("\nTotal Income (%d accounts):\n\n\twtfskins:  $%.2f\n\tcsgolives: $%.2f\n\tpvpro:     $%.2f (%d coins)\n\nOverall:   $%.2f\n",
		len(accounts), totalWtfskinsIncome, totalCsgolivesIncome, totalPvproDollarsIncome, totalPvproCoinsIncome, totalOverallIncomeInDollars)

	/*for {
		fmt.Print("\n\nDo you want to store values to .xls file? (y/n): ")
		var userInput string
		fmt.Scanln(&userInput)

		if userInput == "y" {
			break
		} else if userInput == "n" {
			fmt.Printf("\n\nPress any key to exit ...")
			fmt.Scanln()
			return
		} else {
			continue
		}
	}*/

	//Write To fileMonth income for the whole month

	incomeSheetName := exFile.GetSheetName(0)

	for i := 'B'; i <= 'D'; i++ {
		for j := range accounts {

			if i == 'B' {
				err := exFile.SetCellValue(incomeSheetName, "B"+strconv.Itoa(j+2),
					fmt.Sprintf("+$%.2f", accounts[j].wtfskinsIncome))

				if err != nil {
					fmt.Println(err)
				}
			} else if i == 'C' {
				err := exFile.SetCellValue(incomeSheetName, "C"+strconv.Itoa(j+2),
					fmt.Sprintf("+$%.2f", accounts[j].csgoliveIncome))

				if err != nil {
					fmt.Println(err)
				}
			} else if i == 'D' {
				err := exFile.SetCellValue(incomeSheetName, "D"+strconv.Itoa(j+2),
					"+"+strconv.Itoa(accounts[j].pvproCoinsIncome)+" coins")

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
		color.Green("\n\nCalculated values have been successfully stored to .xls file")
	}

	//Work with IncomeFile

	//2021年12月のルーレット.xlsx
	//fileMonthNameS := fileMonthName.Text()
	fileMonthNameS := "2021年12月のルーレット.xlsx"
	incomeYear, _ := strconv.Atoi(fileMonthNameS[:4])
	incomeMonth, _ := strconv.Atoi(fileMonthNameS[7:strings.Index(fileMonthNameS, "月")]) //7 cuz 年 is 3 bytes

	exFileIncome, err := excelize.OpenFile(workFolder + "\\" + fileTotalIncomeName)

	if err != nil {
		fmt.Println(err)
		fmt.Scanln()
		return
	}

	sheetsNumber := 0
	for {
		incomeSheetName = exFileIncome.GetSheetName(sheetsNumber)

		if incomeSheetName == "" {
			break
		}

		sheetsNumber++
	}

	sheetYearIndex := incomeYear - 2020 //2020 - first sheet (at index 0)

	resp, err := http.Get("https://api.privatbank.ua/p24api/exchange_rates?json&date=" + time.Now().Format("02.01.2006"))

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

	//"baseCurrency":"UAH","currency":"USD" "saleRate":27.4500000,"purchaseRate":27.0500000},
	type Currency struct {
		ForeignCurrencyName string  `json:"currency"`
		PurchaseRate        float64 `json:"purchaseRate"`
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
			totalOverallIncomeInHryvnia = currency.PurchaseRate * totalOverallIncomeInDollars
		}
	}

	//2 loops cuz we need to get income in hryvnia first
	for _, currency := range currencies {
		if currency.ForeignCurrencyName == "RUB" {
			totalOverallIncomeInRubles = totalOverallIncomeInHryvnia / currency.PurchaseRate
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

	err = exFileIncome.Save()

	if err != nil {
		fmt.Println(err)
		fmt.Scanln()
		return
	}

	//fmt.Printf("\n\nPress any key to exit ...")
	//fmt.Scanln()
}
