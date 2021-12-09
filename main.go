package main

import (
	"fmt"
	"net/http"
	"strconv"
	"strings"
	"time"

	"github.com/fatih/color"
	"github.com/xuri/excelize/v2"
)

//Проверять есть ли значение в клетке - и если оно есть и не такое же точно - уведомить об этом, если одинаковы то не записывать
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
		{login: "l///"},
		{login: "r....."},
		{login: "d....."},
		{login: "d....."},
		{login: "d...."},
	}

	var totalWtfskinsIncome, totalCsgolivesIncome, totalPvproDollarsIncome float64
	var totalPvproCoinsIncome int
	var totalOverallIncome float64

	//both functions store data to the struct
	checkAndGetFirstCells(exFile, &accounts)
	getLastCellValues(exFile, &accounts)

	//Print income of each account and count the total income in loop to print it later
	for i := range accounts {
		switch accounts[i].login { //index is needed cuz range-loop copies accounts[i] to account, but not a pointer
		case "l.....":
			color.Red(accounts[i].CalculateS())
		case "r.....":
			color.Magenta(accounts[i].CalculateS())
		case "d.....":
			color.Yellow(accounts[i].CalculateS())
		case ".....":
			color.Cyan(accounts[i].CalculateS())
		case "d.....":
			color.White(accounts[i].CalculateS())
		}

		totalWtfskinsIncome += accounts[i].wtfskinsIncome
		totalCsgolivesIncome += accounts[i].csgoliveIncome

		totalPvproCoinsIncome += accounts[i].pvproCoinsIncome
		totalPvproDollarsIncome += accounts[i].pvproDollarsIncome

	}

	//print Total Income
	totalOverallIncome = totalWtfskinsIncome + totalCsgolivesIncome + totalPvproDollarsIncome
	color.Green("\nTotal Income (%d accounts):\n\n\twtfskins:  $%.2f\n\tcsgolives: $%.2f\n\tpvpro:     $%.2f (%d coins)\n\nOverall:   $%.2f\n",
		len(accounts), totalWtfskinsIncome, totalCsgolivesIncome, totalPvproDollarsIncome, totalPvproCoinsIncome, totalOverallIncome)

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

	if err := exFile.SetCellValue(incomeSheetName, "C11", fmt.Sprintf("Total Income: $%.2f", totalOverallIncome)); err != nil {
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

	buff := make([]byte, 5000) //5000 characters should be enough for Privat24 Json response
	resp.Body.Read(buff)

	//remove useless from the beginning ""date":"09.12.2021","bank":"PB","baseCurrency":980,"baseCurrencyLit":"UAH","exchangeRate":["
	privatDataS := string(buff)[strings.Index(string(buff), "[")+1:]
	fmt.Println(privatDataS)

	err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "B"+strconv.Itoa(incomeMonth+1),
		fmt.Sprintf("$%.2f", totalOverallIncome))

	if err != nil {
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
