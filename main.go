package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"

	"github.com/fatih/color"
	"github.com/xuri/excelize/v2"
)

func main() {

	workFolder := "D:\\Program Files\\...."

	fmt.Print(workFolder + ": ")

	fileName := bufio.NewScanner(os.Stdin) //2021年12月のルーレット.xlsx
	fileName.Scan()

	if fileName.Err() != nil {
		fmt.Println(fileName.Err())
		fmt.Scanln()
		return
	}

	exFile, err := excelize.OpenFile(workFolder + "\\" + fileName.Text())

	if err != nil {
		fmt.Println(err)
		fmt.Scanln()
		return
	}

	accounts := []Accounts{
		{login: "l....."},
		{login: "r......"},
		{login: "d......."},
		{login: ".....1"},
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
		case ".......":
			color.Red(accounts[i].CalculateS())
		case "......":
			color.Magenta(accounts[i].CalculateS())
		case "d...":
			color.Yellow(accounts[i].CalculateS())
		case "d.":
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

	for {
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
	}

	//Write To Excel block

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

	fmt.Printf("\n\nPress any key to exit ...")
	fmt.Scanln()
}
