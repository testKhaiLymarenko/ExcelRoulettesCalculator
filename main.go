package main

import (
	"fmt"
	"os"
	"os/exec"
	"runtime"
	"strconv"
	"strings"
	"time"

	"github.com/fatih/color"
)

//add Linux support
//add Linux WSL support
//файлы-параметры (много, если какой-то ошибочный - вывод об ошибки и продолжить работу)
//fix error color.Red vo vsex filax

func main() {
	if runtime.GOOS == "windows" { //change console window title
		cmd := exec.Command("cmd", "/C", "title", "ExcelRoulettesCalculator")
		if err := cmd.Run(); err != nil {
			color.Red("%v", err)
		}
	}

	defer func() {
		fmt.Printf("\nPress any key to exit ...")
		fmt.Scanln()
	}()

	////"D:\\Program Files\\MEGAsync\\MEGAsync\\Internet Deals\\Steam\\ルーレット"
	workDir := "D:\\Program Files\\MEGAsync\\MEGAsync\\Internet Deals\\Steam\\ルーレット" //"D:\\Program Files\\MEGA\\Internet Deals\\Steam\\ルーレット"
	var workDirExists bool
	currDir, _ := os.Getwd()
	currDir = strings.ToUpper(currDir[:1]) + currDir[1:] // (f:\\ -> F:\\)
	fileTotalIncomeName := "_ルーレットの総収入.xlsx"

	printStartMessage(workDir, currDir, &workDirExists)

	/*fileMonthName := bufio.NewScanner(os.Stdin) //2021年12月のルーレット.xlsx
	fileMonthName.Scan()

	if fileMonthName.Err() != nil {
		fmt.Println(fileMonthName.Err())
		fmt.Scanln()
		return
	}
	*/

	fileMonthName := "2022年1月のルーレット.xlsx" // --> for debug
	exFile, err := getExcelFileHandle(workDirExists, workDir, currDir, fileMonthName)
	if err != nil {
		color.Red("%v", err)
		return
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
	if err := checkAndGetFirstCells(exFile, &accounts); err != nil {
		return //error is shown to console straight from the function
	}
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
	writeToFileOverallIncome(workDir, fileTotalIncomeName, fileMonthName, totalOverallIncomeInDollars, incomeMonth, incomeYear, fileMonthNameAlias)
}

func printStartMessage(workDir, currDir string, workDirExists *bool) {
	if !fileExists(workDir) {
		nonExistedFile, err := nonExistedFirstDir(workDir)
		if err != nil {
			color.Red("%v", err)
			return
		}

		if nonExistedFile != workDir {
			color.Red("Working directory: "+workDir+"\\ not found because %s is missing", nonExistedFile)
		} else {
			color.Red("Working directory: " + workDir + "\\ not found")
		}

		fmt.Print("Current directory: " + currDir + "\\: ")

		*workDirExists = false
	} else {
		fmt.Println("Current directory: " + currDir + "\\")
		fmt.Print("Working directory: " + workDir + "\\")
		*workDirExists = true
	}

	fmt.Println()
}

//Check if file or directory exists
func fileExists(path string) bool {
	if _, err := os.Stat(path); err == nil {
		return true
	}

	return false
}
