package main

import (
	"fmt"
	"os"
	"runtime"
	"strconv"
	"strings"
	"time"

	"github.com/fatih/color"
	"github.com/xuri/excelize/v2"
)

//добавить название для папок в вывод WorkingDirectory: path: ...
//файлы-параметры (много, если какой-то ошибочный - вывод об ошибки и продолжить работу)
//os.Exit v check..()

func main() {
	defer func() {
		fmt.Printf("\nPress any key to exit ...")
		fmt.Scanln()
	}()

	workFolder := "D:\\Program Files\\MEGA\\Internet Deals\\Steam\\ルーレット"
	var workFolderExists bool
	currFolder, _ := os.Getwd()
	fileTotalIncomeName := "_ルーレットの総収入.xlsx"

	if !fileExists(workFolder) {
		nonExistedFile, err := nonExistedFirstDir(workFolder)
		if err != nil {
			color.Red("%v", err)
			return
		}

		if nonExistedFile != workFolder {
			color.Red(workFolder+"\\ not found because %s doesn't exist", nonExistedFile)
		} else {
			color.Red(workFolder + "\\ not found")
		}

		color.Yellow(currFolder + ": ")

		workFolderExists = false
	} else {
		color.Yellow(currFolder + "\\: ")
		color.Yellow(workFolder + "\\: ")
		workFolderExists = true
	}

	fmt.Println()
	/*fileMonthName := bufio.NewScanner(os.Stdin) //2021年12月のルーレット.xlsx
	fileMonthName.Scan()

	if fileMonthName.Err() != nil {
		fmt.Println(fileMonthName.Err())
		fmt.Scanln()
		return
	}
	*/

	fileMonthName := "2022年1月のルーレット.xlsx" // --> for debug
	var exFile *excelize.File
	var err error

	//добавить стирание консоли если файл найден в основной папке
	if workFolderExists && fileExists(workFolder+"\\"+fileMonthName) {
		//exFile, err := excelize.OpenFile(workFolder + "\\" + fileMonthName.Text())
		exFile, err = excelize.OpenFile(workFolder + "\\" + fileMonthName) // --> for debug
		if err != nil {
			color.Red("%v", err)
			return
		}
	} else {
		if runtime.GOOS == "windows" {
			exFile, err = excelize.OpenFile(currFolder + "\\" + fileMonthName) //"2021年12月のルーレット.xlsx")
		} else {
			exFile, err = excelize.OpenFile(currFolder + "/" + fileMonthName) // "2021年12月のルーレット.xlsx")
		}

		if err != nil {
			color.Red("%v\n\n", err)
			return
		}
	}

	//Get account names from excel file but not hardcoded
	var accountNames []string

	sheetsNumber := 0
	for {
		//first sheet is the month
		if exFile.GetSheetName(sheetsNumber) != "" {
			if sheetsNumber > 0 {
				accountNames = append(accountNames, exFile.GetSheetName(sheetsNumber))
			}
		} else {
			break
		}

		sheetsNumber++
	}

	accounts := make([]Accounts, len(accountNames))

	for i := 0; i < len(accountNames); i++ {
		accounts[i].login = accountNames[i]
	}

	accounts = make([]Accounts, len(accountNames))

	for i := 0; i < len(accountNames); i++ {
		accounts[i].login = accountNames[i]
	}

	accounts = make([]Accounts, len(accountNames))

	for i := 0; i < len(accountNames); i++ {
		accounts[i].login = accountNames[i]
	}

	var totalWtfskinsIncome, totalCsgolivesIncome, totalPvproDollarsIncome float64
	var totalPvproCoinsIncome int
	var totalOverallIncomeInDollars float64

	//both functions store data to the struct
	checkAndGetFirstCells(exFile, &accounts)
	getLastCellValues(exFile, &accounts)

	//Print income of each account and count the total income in loop to print it later
	fmt.Println()
	for i := range accounts {
		switch i { //index is needed cuz range-loop copies accounts[i] to account, but not a pointer
		case 0:
			color.Red(accounts[i].CalculateS())
		case 1:
			color.Magenta(accounts[i].CalculateS())
		case 2:
			color.Yellow(accounts[i].CalculateS())
		case 3:
			color.Cyan(accounts[i].CalculateS())
		case 4:
			color.White(accounts[i].CalculateS())
		default:
			color.Yellow(accounts[i].CalculateS())
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

	//!!!!writeExcel section

	//All of these needed just to print Month_year.xlsx (December_2021.xlsx)
	//2021年12月のルーレット.xlsx
	incomeMonth, _ := strconv.Atoi(fileMonthName[7:strings.Index(fileMonthName, "月")]) //7 cuz '年' consists of 3 bytes//7 cuz '年' consists of 3 bytes
	incomeYear, _ := strconv.Atoi(fileMonthName[:4])
	monthT := time.Month(incomeMonth)
	fileMonthNameAlias := monthT.String() + "_" + strconv.Itoa(incomeYear) + ".xlsx"

	//1st file
	writeToFileMonthIncome(exFile, &accounts, totalWtfskinsIncome, totalCsgolivesIncome, totalPvproDollarsIncome,
		totalOverallIncomeInDollars, totalPvproCoinsIncome, fileMonthName, fileMonthNameAlias)

	//2nd file
	writeToFileOverallIncome(workFolder, fileTotalIncomeName, fileMonthName, totalOverallIncomeInDollars, incomeMonth, incomeYear, fileMonthNameAlias)
}
