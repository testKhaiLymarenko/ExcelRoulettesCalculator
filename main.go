package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"os"
	"runtime"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/fatih/color"
	"github.com/xuri/excelize/v2"
)

//файлы не только в меге но и рядом с файлом - для работы на линуксе и андроиде

func main() {

	workFolder := "D:\\Program Files\\MEGAsync\\MEGAsync\\Internet Deals\\Steam\\ルーレット"
	currFolder, _ := os.Getwd()
	fileTotalIncomeName := "_ルーレットの総収入.xlsx"

	fmt.Println(currFolder + ": ")
	fmt.Print(workFolder + ": ")

	/*fileMonthName := bufio.NewScanner(os.Stdin) //2021年12月のルーレット.xlsx
	fileMonthName.Scan()

	if fileMonthName.Err() != nil {
		fmt.Println(fileMonthName.Err())
		fmt.Scanln()
		return
	}


	exFile, err := excelize.OpenFile(workFolder + "\\" + fileMonthName.Text())*/

	fileMonthName := "2021年12月のルーレット.xlsx"
	exFile, err := excelize.OpenFile(workFolder + "\\" + "2021年12月のルーレット.xlsx") // --> for debug

	if err != nil {
		var err2 error

		if runtime.GOOS == "windows" {
			exFile, err2 = excelize.OpenFile(currFolder + "\\" + "2021年12月のルーレット.xlsx")
		} else {
			exFile, err2 = excelize.OpenFile(currFolder + "/" + "2021年12月のルーレット.xlsx")
		}

		if err2 != nil {
			fmt.Printf("%s\n%s\n\n", err, err2)
			fmt.Printf("\n\nPress any key to exit ...")
			fmt.Scanln()
			return
		}
	}

	accounts := []Accounts{
		{login: "."},
		{login: "ra."},
		{login: "d."},
		{login: "d.."},
		{login: "d.."},
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
		case "...":
			color.Red(accounts[i].CalculateS())
		case "..":
			color.Magenta(accounts[i].CalculateS())
		case "de..":
			color.Yellow(accounts[i].CalculateS())
		case "d.1":
			color.Cyan(accounts[i].CalculateS())
		case ".":
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

	// writeExcel section

	//All of these needed just to print Month_year.xlsx (December_2021.xlsx)
	//2021年12月のルーレット.xlsx
	incomeMonth, _ := strconv.Atoi(fileMonthName[7:strings.Index(fileMonthName, "月")]) //7 cuz '年' consists of 3 bytes//7 cuz '年' consists of 3 bytes
	incomeYear, _ := strconv.Atoi(fileMonthName[:4])
	monthT := time.Month(incomeMonth)
	fileMonthNameAlias := monthT.String() + "_" + strconv.Itoa(incomeYear) + ".xlsx"

	//Ask (1/2)
	/*for {
		fmt.Printf("\n\nDo you want to store values into %s? (y/n): ", fileMonthNameAlias)
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

	writeToFileMonthIncome(exFile, &accounts, totalWtfskinsIncome, totalCsgolivesIncome, totalPvproDollarsIncome,
		totalOverallIncomeInDollars, totalPvproCoinsIncome, fileMonthName, fileMonthNameAlias)

	//Ask (2/2)
	fileMonthNameAlias = "Total_Income.xlsx"
	/*for {
		fmt.Print("\n\nDo you want to store values into %s? (y/n): ", fileMonthNameAlias)
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

	writeToFileOverallIncome(workFolder, fileTotalIncomeName, fileMonthName, totalOverallIncomeInDollars, incomeMonth, incomeYear, fileMonthNameAlias)

	fmt.Printf("\n\nPress any key to exit ...")
	fmt.Scanln()
}

type Accounts struct {
	login string

	wtfskinsFirstVal float64 //dollars
	wtfskinsLastVal  float64
	csgoliveFirstVal float64
	csgoliveLastVal  float64
	pvproFirstVal    int //coins
	pvproLastVal     int

	wtfskinsIncome     float64
	csgoliveIncome     float64
	pvproCoinsIncome   int
	pvproDollarsIncome float64
}

func (a *Accounts) CalculateS() string {

	//last value has 0 only if every cell has no value, so income values have 0 by default

	if a.wtfskinsLastVal != 0 {
		a.wtfskinsIncome = a.wtfskinsLastVal - a.wtfskinsFirstVal
	}

	if a.csgoliveLastVal != 0 {
		a.csgoliveIncome = a.csgoliveLastVal - a.csgoliveFirstVal
	}

	if a.pvproLastVal != 0 {
		a.pvproCoinsIncome = a.pvproLastVal - a.pvproFirstVal
		a.pvproDollarsIncome = float64(a.pvproCoinsIncome) / 1000
	}

	return "\t\t\t" + a.login + ":\nwtfskins: " + fmt.Sprintf("$%.2f\n", a.wtfskinsIncome) + "csgolive: " +
		fmt.Sprintf("$%.2f\n", a.csgoliveIncome) + "pvpro:    $" + fmt.Sprintf("%.2f (%d coins)\n\n", a.pvproDollarsIncome, a.pvproCoinsIncome)
}

//gets last value and add cashout cell values to them
func getLastCellValues(exFile *excelize.File, accounts *[]Accounts) {

	var wg sync.WaitGroup

	//get a value of the last cell that contains it (may be not B31, B32, etc but can be whatever B5, B20 - dynamic)
	getLastCellValue := func(account *Accounts) {

		//B - wtfskins, C - csgolive, D - pvpro
		for i := 'B'; i <= 'D'; i++ {
			for j := 33; j > 2; j-- {
				cell, _ := exFile.GetCellValue(account.login, string(i)+strconv.Itoa(j))
				if cell != "" {

					if i == 'B' || i == 'C' { //wtfskins and csgolive

						if val, err := strconv.ParseFloat(cell[1:], 64); err != nil {
							continue
						} else {
							if i == 'B' {
								account.wtfskinsLastVal = val
							} else if i == 'C' {
								account.csgoliveLastVal = val
							}
							break
						}

					} else if i == 'D' { //pvpro

						//4433 coins + 560 crystals
						if val, err := strconv.Atoi(cell[:strings.Index(cell, " ")]); err != nil {
							continue
						} else {
							account.pvproLastVal = val
							break
						}

					}

				}

			}

		}
		wg.Done() //tells that this goroutine finished working
	}

	//check if there were cashout in this month i.e. ($0.29 -> $0.02) and if yes - add the substraction !!!!!to the last value!!!!
	addCashoutCellValues := func(account *Accounts) {

		//B - wtfskins, C - csgolive, D - pvpro
		for i := 'B'; i <= 'D'; i++ {
			for j := 3; j <= 33; j++ {
				cell, _ := exFile.GetCellValue(account.login, string(i)+strconv.Itoa(j))
				if strings.Contains(cell, "->") {
					if i == 'B' || i == 'C' { //Example: $0.29 -> $0.02
						before, _ := strconv.ParseFloat(cell[1:strings.Index(cell, " ->")], 64)
						after, _ := strconv.ParseFloat(cell[strings.Index(cell, " $")+2:], 64)

						if i == 'B' {
							account.wtfskinsLastVal += (before - after)
						} else if i == 'C' {
							account.csgoliveLastVal += (before - after)
						}
					} else if i == 'D' { //Example: 4088 coins + 190 crystals -> 3989 coins + 190 crystals
						before, _ := strconv.Atoi(cell[:strings.Index(cell, " ")])
						after, _ := strconv.Atoi(cell[strings.Index(cell, "> ")+2 : strings.LastIndex(cell, " co")])

						account.pvproLastVal += (before - after)
					}
				}

			}

		}

		wg.Done() //tells that this goroutine finished working
	}

	for i := range *accounts {
		go getLastCellValue(&(*accounts)[i]) //cuz value is copied but not &
		wg.Add(1)
	}
	wg.Wait()

	for i := range *accounts {
		go addCashoutCellValues(&(*accounts)[i]) //cuz value is copied but not &
		wg.Add(1)
	}
	wg.Wait()
}

//check if first cells contain data from the previous month and fill the accounts' struct if yes
func checkAndGetFirstCells(exFile *excelize.File, accounts *[]Accounts) {

	var log string //collect accounts which don't have value for wtfskins, csgolives, pvpro
	empty := false //need to check if there's at least one account that doesn't have a value for any of 3 roulettes

	//B - wtfskins, C - csgolive, D - pvpro
	for i := 'B'; i <= 'D'; i++ {

		/*
			needed to have error log look this way:
													wtfskins: account1 account2
													csgolive: account3
													pvpro:
		*/
		if i == 'B' {
			log += "wtfskins: "
		} else if i == 'C' {
			log += "csgolive: "
		} else if i == 'D' {
			log += "pvpro:    "
		}

		//check only one column of each account
		for j := 0; j < len(*accounts); j++ {
			login := (*accounts)[j].login

			cell, _ := exFile.GetCellValue(login, string(i)+"2") //B + 2 = B2 cell

			if i == 'B' {
				if cell == "" { //
					empty = true
					log += login + " "
				} else {
					if val, err := strconv.ParseFloat(cell[1:], 64); err == nil {
						(*accounts)[j].wtfskinsFirstVal = val
					} else {
						fmt.Println(err)
					}
				}
			} else if i == 'C' {
				if cell == "" {
					empty = true
					log += login + " "
				} else {
					if val, err := strconv.ParseFloat(cell[1:], 64); err == nil {
						(*accounts)[j].csgoliveFirstVal = val
					} else {
						fmt.Println(err)
					}
				}
			} else if i == 'D' {
				if cell == "" {
					empty = true
					log += login + " "
				} else {
					if val, err := strconv.Atoi(cell[:strings.Index(cell, " ")]); err == nil {
						(*accounts)[j].pvproFirstVal = val
					} else {
						fmt.Println(err)
					}
				}
			}
		}

		log += "\n" //border between wtfskins, csgolive, pvpro
	}

	if empty {
		roulettes := strings.Split(log, "\n")

		color.Red("Error! List of cells that doesn't contain last month data in the first raw:\n\n")
		for i := 0; i < len(roulettes); i++ {
			switch i {
			case 0:
				color.Green(roulettes[i])
			case 1:
				color.Cyan(roulettes[i])
			case 2:
				color.Yellow(roulettes[i])
			}
		}

		fmt.Scanln()
		os.Exit(1)
	}
}

func writeToFileMonthIncome(exFile *excelize.File, accounts *[]Accounts, totalWtfskinsIncome, totalCsgolivesIncome,
	totalPvproDollarsIncome, totalOverallIncomeInDollars float64, totalPvproCoinsIncome int,
	fileMonthName, fileMonthNameAlias string) {
	incomeSheetName := exFile.GetSheetName(0)

	//Write data to account cells: wtfskins, csgolive, pvpro

	//store current cell values in beforeS slice, and new values in afterS
	var beforeS, afterS []string

	for i := 'B'; i <= 'D'; i++ {
		for j := range *accounts {

			if i == 'B' {
				//Check if current and storead values are the same, if they are different - overwrite it here (it will apply)
				//	only if we call Save() and store these values in 2 different slices to show the difference later
				if val, err := exFile.GetCellValue(incomeSheetName, "B"+strconv.Itoa(j+2)); err != nil {
					fmt.Println(err)
				} else {
					accountWtfskinsIncomeS := fmt.Sprintf("+$%.2f", (*accounts)[j].wtfskinsIncome)

					if val != accountWtfskinsIncomeS {
						if err := exFile.SetCellValue(incomeSheetName, "B"+strconv.Itoa(j+2), accountWtfskinsIncomeS); err != nil {
							fmt.Println(err)
						} else {
							beforeS = append(beforeS, "wtfskins: "+val)
							afterS = append(afterS, "wtfskins: "+accountWtfskinsIncomeS)
						}
					}

				}
			} else if i == 'C' {
				if val, err := exFile.GetCellValue(incomeSheetName, "C"+strconv.Itoa(j+2)); err != nil {
					fmt.Println(err)
				} else {
					accountCsgoliveIncomeS := fmt.Sprintf("+$%.2f", (*accounts)[j].csgoliveIncome)

					if val != accountCsgoliveIncomeS {
						if err := exFile.SetCellValue(incomeSheetName, "C"+strconv.Itoa(j+2), accountCsgoliveIncomeS); err != nil {
							fmt.Println(err)
						} else {
							beforeS = append(beforeS, "csgolive: "+val)
							afterS = append(afterS, "csgolive: "+accountCsgoliveIncomeS)
						}
					}

				}
			} else if i == 'D' {
				if val, err := exFile.GetCellValue(incomeSheetName, "D"+strconv.Itoa(j+2)); err != nil {
					fmt.Println(err)
				} else {
					accountPvproIncomeCoinsS := fmt.Sprintf("+%d coins", (*accounts)[j].pvproCoinsIncome)

					if val != accountPvproIncomeCoinsS {
						if err := exFile.SetCellValue(incomeSheetName, "D"+strconv.Itoa(j+2), accountPvproIncomeCoinsS); err != nil {
							fmt.Println(err)
						} else {
							beforeS = append(beforeS, "pvpro: "+val)
							afterS = append(afterS, "pvpro: "+accountPvproIncomeCoinsS)
						}
					}

				}
			}
		}
	}

	//if current and new values are different
	if len(beforeS) > 0 {

		fmt.Print("\n\n")

		//account1: [0], [1], [3] | account2: [4], [5], [6] | ....
		beforeSlice := make([][]string, len(*accounts))
		afterSlice := make([][]string, len(*accounts))
		//index for loop in accounts names
		var j int

		for i := 0; i < len(beforeS); i++ {

			//wtfskins, csgolive and pvpro values changes every N elements
			if i%len(*accounts) == 0 {
				j = 0
			}

			beforeSlice[j] = append(beforeSlice[j], beforeS[i])
			afterSlice[j] = append(afterSlice[j], afterS[i])

			j++
		}

		//print values in just one line
		for i := 0; i < len(*accounts); i++ {
			color.Yellow((*accounts)[i].login + ": ")
			color.Red("\t" + strings.Join(beforeSlice[i], "\t"))
			color.Cyan("\t" + strings.Join(afterSlice[i], "\t"))
		}
	}

	// Write data to OVERALL cells

	//reset slices to use it for OVERALL cells
	beforeS = nil
	afterS = nil

	//Check if current and storead values are the same, if they are different - overwrite it here (it will apply)
	//	only if we call Save() and store these values in 2 different slices to show the difference later

	if val, err := exFile.GetCellValue(incomeSheetName, "B8"); err != nil {
		fmt.Println(err)
	} else {
		totalWtfskinsIncomeS := fmt.Sprintf("+$%.2f", totalWtfskinsIncome)
		if val != totalWtfskinsIncomeS {
			if err := exFile.SetCellValue(incomeSheetName, "B8", totalWtfskinsIncomeS); err != nil {
				fmt.Println(err)
			} else {
				beforeS = append(beforeS, "\twtfskins: "+val)
				afterS = append(afterS, "\twtfskins: "+totalWtfskinsIncomeS)
			}
		}
	}

	if val, err := exFile.GetCellValue(incomeSheetName, "C8"); err != nil {
		fmt.Println(err)
	} else {
		totalCsgolivesIncomeS := fmt.Sprintf("+$%.2f", totalCsgolivesIncome)
		if val != totalCsgolivesIncomeS {
			if err := exFile.SetCellValue(incomeSheetName, "C8", totalCsgolivesIncomeS); err != nil {
				fmt.Println(err)
			} else {
				beforeS = append(beforeS, "\tcsgolive: "+val)
				afterS = append(afterS, "\tcsgolive: "+totalCsgolivesIncomeS)
			}
		}
	}

	if val, err := exFile.GetCellValue(incomeSheetName, "D8"); err != nil {
		fmt.Println(err)
	} else {
		totalPvproIncomeS := fmt.Sprintf("+%d coins (+$%.2f)", totalPvproCoinsIncome, totalPvproDollarsIncome)
		if val != totalPvproIncomeS {
			if err := exFile.SetCellValue(incomeSheetName, "D8", totalPvproIncomeS); err != nil {
				fmt.Println(err)
			} else {
				beforeS = append(beforeS, "\tpvpro: "+val)
				afterS = append(afterS, "\tpvpro: "+totalPvproIncomeS)
			}
		}
	}

	if val, err := exFile.GetCellValue(incomeSheetName, "C11"); err != nil {
		fmt.Println(err)
	} else {
		totalOverallIncomeInDollarsS := fmt.Sprintf("Total Income: $%.2f", totalOverallIncomeInDollars)
		if val != totalOverallIncomeInDollarsS {
			if err := exFile.SetCellValue(incomeSheetName, "C11", totalOverallIncomeInDollarsS); err != nil {
				fmt.Println(err)
			} else {
				beforeS = append(beforeS, "\t"+val)
				afterS = append(afterS, "\t"+totalOverallIncomeInDollarsS)
			}
		}
	}

	if len(beforeS) > 0 {
		color.Yellow("\n\nOVERALL: ")
		color.Red(strings.Join(beforeS, "  "))
		color.Cyan(strings.Join(afterS, "  "))

		//Ask user if user wants to overwrite data

		for {
			fmt.Printf("\n\nDo you want to overwrite values in %s? (y/n): ", fileMonthNameAlias)
			var userInput string
			fmt.Scanln(&userInput)

			if userInput == "y" {
				break
			} else if userInput == "n" {
				fmt.Println(strings.Repeat("_", 110))
				return
			} else {
				continue
			}
		}

		if err := exFile.Save(); err != nil {
			color.Red("\n%s", err)
			color.Red("\n\nCalculated values have not been stored into %s", fileMonthNameAlias)
		} else {
			color.Green("\n\nCalculated values have been successfully stored into %s", fileMonthNameAlias)
		}
	}

	fmt.Println(strings.Repeat("_", 110))
}

func writeToFileOverallIncome(workFolder, fileTotalIncomeName, fileMonthName string, totalOverallIncomeInDollars float64,
	incomeMonth, incomeYear int, fileMonthNameAlias string) {

	fmt.Println()

	var (
		exFileIncome *excelize.File
		err          error
	)

	exFileIncome, err = excelize.OpenFile(workFolder + "\\" + fileTotalIncomeName)
	if err != nil {
		var err2 error

		currFolder, _ := os.Getwd()
		if runtime.GOOS == "windows" {
			exFileIncome, err2 = excelize.OpenFile(currFolder + "\\" + fileTotalIncomeName)
		} else {
			exFileIncome, err2 = excelize.OpenFile(currFolder + "/" + fileTotalIncomeName)
		}

		if err2 != nil {
			fmt.Printf("%s\n%s\n\n", err, err2)
			fmt.Printf("\n\nPress any key to exit ...")
			fmt.Scanln()
			return
		}
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

	var hryvniaRate, rubleRate, totalOverallIncomeInHryvnias, totalOverallIncomeInRubles float64

	for _, currency := range currencies {
		if currency.ForeignCurrencyName == "USD" {
			hryvniaRate = currency.PurchaseRateNBU //how many hryvnias in $1
			totalOverallIncomeInHryvnias = hryvniaRate * totalOverallIncomeInDollars
		}
	}

	//2 loops cuz we need to get income in hryvnia first
	for _, currency := range currencies {
		if currency.ForeignCurrencyName == "RUB" {
			rubleRate = currency.PurchaseRateNBU
			totalOverallIncomeInRubles = totalOverallIncomeInHryvnias / rubleRate
		}
	}

	var beforeS, afterS []string

	color.Yellow("Income ($, ₽, ₴):  $1 = ₴%.2f = ₽%.2f [НБУ, %s]", hryvniaRate, hryvniaRate/rubleRate, date)

	//Set value for the Dollars cell
	if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "B"+strconv.Itoa(incomeMonth+1)); err != nil {
		fmt.Println(err)
	} else {
		totalOverallIncomeInDollarsS := fmt.Sprintf("$%.2f", totalOverallIncomeInDollars)
		if val != totalOverallIncomeInDollarsS {
			if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "B"+strconv.Itoa(incomeMonth+1),
				totalOverallIncomeInDollarsS); err != nil {
				fmt.Println(err)
				return
			} else {
				beforeS = append(beforeS, "\t"+val)
				afterS = append(afterS, "\t"+totalOverallIncomeInDollarsS)
			}
		}
	}

	//Set value for the Rubles cell
	if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "C"+strconv.Itoa(incomeMonth+1)); err != nil {
		fmt.Println(err)
	} else {
		totalOverallIncomeInRublesS := fmt.Sprintf("₽%.2f", totalOverallIncomeInRubles)
		if val != totalOverallIncomeInRublesS {
			if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "C"+strconv.Itoa(incomeMonth+1),
				totalOverallIncomeInRublesS); err != nil {
				fmt.Println(err)
				return
			} else {
				beforeS = append(beforeS, "\t"+val)
				afterS = append(afterS, "\t"+totalOverallIncomeInRublesS)
			}
		}
	}

	//Set value for the Hryvnia cell
	if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "D"+strconv.Itoa(incomeMonth+1)); err != nil {
		fmt.Println(err)
	} else {
		totalOverallIncomeInHryvniasS := fmt.Sprintf("₴%.2f", totalOverallIncomeInHryvnias)
		if val != totalOverallIncomeInHryvniasS {
			if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "D"+strconv.Itoa(incomeMonth+1),
				totalOverallIncomeInHryvniasS); err != nil {
				fmt.Println(err)
				return
			} else {
				beforeS = append(beforeS, "\t"+val)
				afterS = append(afterS, "\t"+totalOverallIncomeInHryvniasS)
			}
		}
	}

	// Calculate all the year

	var yearIncomeInDollars, yearIncomeInRubles, yearIncomeInHryvnias float64

	for i := 'B'; i <= 'D'; i++ {
		for j := 2; j <= 13; j++ {
			if i == 'B' {
				if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "B"+strconv.Itoa(j)); err != nil {
					fmt.Println(err)
				} else {
					if len(val) > 0 { // in case of an empty string
						if num, err := strconv.ParseFloat(val[1:], 64); err == nil { //$ = 1 byte ($11.80)
							yearIncomeInDollars += num
						}
					}
				}
			} else if i == 'C' {
				if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "C"+strconv.Itoa(j)); err != nil {
					fmt.Println(err)
				} else {
					if len(val) > 0 {
						if num, err := strconv.ParseFloat(val[3:], 64); err == nil { //₽ = 3 bytes (₽881.92)
							yearIncomeInRubles += num
						}
					}
				}
			} else if i == 'D' {
				if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "D"+strconv.Itoa(j)); err != nil {
					fmt.Println(err)
				} else {
					if len(val) > 0 {
						if num, err := strconv.ParseFloat(val[3:], 64); err == nil { //₴ = 3 bytes (₴318.90)
							yearIncomeInHryvnias += num
						}
					}
				}
			}
		}
	}

	if len(beforeS) > 0 {
		color.Red(strings.Join(beforeS, "\t"))
		color.Cyan(strings.Join(afterS, "\t"))
	}

	beforeS = nil
	afterS = nil

	//Set value for the YearDollars cell
	if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "B14"); err != nil {
		fmt.Println(err)
	} else {
		yearIncomeInDollarsS := fmt.Sprintf("$%.2f", yearIncomeInDollars)
		if val != yearIncomeInDollarsS {
			if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "B14", yearIncomeInDollarsS); err != nil {
				fmt.Println(err)
				return
			} else {
				beforeS = append(beforeS, val)
				afterS = append(afterS, yearIncomeInDollarsS)
			}
		}
	}

	//Set value for the YearRubles cell
	if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "C14"); err != nil {
		fmt.Println(err)
	} else {
		yearIncomeInRublesS := fmt.Sprintf("₽%.2f", yearIncomeInRubles)
		if val != yearIncomeInRublesS {
			if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "C14", yearIncomeInRublesS); err != nil {
				fmt.Println(err)
				return
			} else {
				beforeS = append(beforeS, val)
				afterS = append(afterS, yearIncomeInRublesS)
			}
		}
	}

	//Set value for the YearHryvnia cell
	if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "D14"); err != nil {
		fmt.Println(err)
	} else {
		yearIncomeInHryvniasS := fmt.Sprintf("₴%.2f", yearIncomeInHryvnias)
		if val != yearIncomeInHryvniasS {
			if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "D14", yearIncomeInHryvnias); err != nil {
				fmt.Println(err)
				return
			} else {
				beforeS = append(beforeS, val)
				afterS = append(afterS, yearIncomeInHryvniasS)
			}
		}
	}

	fmt.Print("\n\n")
	color.Yellow("Annual income ($, ₽, ₴):")
	if len(beforeS) > 0 {
		color.Red("\t" + strings.Join(beforeS, "  \t"))
		color.Cyan("\t" + strings.Join(afterS, "  \t"))
	}

	for {
		fmt.Printf("\n\nDo you want to overwrite values in %s? (y/n): ", fileMonthNameAlias)
		var userInput string
		fmt.Scanln(&userInput)

		if userInput == "y" {
			break
		} else if userInput == "n" {
			fmt.Println(strings.Repeat("_", 110))
			return
		} else {
			continue
		}
	}

	if err := exFileIncome.Save(); err != nil {
		color.Red("\n%s", err)
		color.Red("\n\nCalculated values have not been stored into %s", fileMonthNameAlias)
	} else {
		color.Green("\n\nCalculated values have been successfully stored into %s", fileMonthNameAlias)
	}

	fmt.Println(strings.Repeat("_", 110))
}
