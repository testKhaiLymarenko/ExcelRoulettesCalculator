package main

import (
	"fmt"

	"github.com/fatih/color"
	"github.com/xuri/excelize/v2"
)

func main() {
	exFile, err := excelize.OpenFile("d:\\list.xlsx")

	if err != nil {
		fmt.Println(err)
		fmt.Scanln()
		return
	}

	accounts := []Accounts{
		{login: "..."},
		{login: "..."},
		{login: "..."},
		{login: "d..."},
		{login: "d..."},
	}

	var totalWtfskinsIncome, totalCsgolivesIncome, totalPvproDollarsIncome float64
	var totalPvproCoinsIncome int

	checkAndGetFirstCells(exFile, &accounts)
	getLastCellName(exFile, &accounts)

	//Print income of each account and count the total income in loop to print it later
	for _, account := range accounts {
		switch account.login {
		case "....":
			color.Red(account.CalculateS())
		case ".....":
			color.Magenta(account.CalculateS())
		case "d...":
			color.Yellow(account.CalculateS())
		case "..":
			color.Cyan(account.CalculateS())
		case "d.":
			color.White(account.CalculateS())
		}

		totalWtfskinsIncome += account.wtfskinsIncome
		totalCsgolivesIncome += account.csgoliveIncome

		totalPvproCoinsIncome += account.pvproCoinsIncome
		totalPvproDollarsIncome += account.pvproDollarsIncome

	}

	//print Total Income
	color.Green("\nTotal Income (%d accounts):\n\n\twtfskins:  $%.2f\n\tcsgolives: $%.2f\n\tpvpro:     $%.2f (%d coins)\n\nOverall:   $%.2f\n",
		len(accounts), totalWtfskinsIncome, totalCsgolivesIncome, totalPvproDollarsIncome, totalPvproCoinsIncome,
		(totalWtfskinsIncome + totalCsgolivesIncome + totalPvproDollarsIncome))

	fmt.Scanln()
}
