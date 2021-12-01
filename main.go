package main

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	exFile, err := excelize.OpenFile("d:\\list.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	IS_PVPRO, NOT_PVPRO := true, false
	logins := [...]string{"...", ".", "d..", "de..", "d."}
	lastCell, err := getLastCell(exFile)

	if err != nil {
		fmt.Println(err)
		return
	}

	roulette := map[string]string{
		"B2": "B" + strconv.Itoa(lastCell), //wtfskins
		"C2": "C" + strconv.Itoa(lastCell), //csgolines
		"D2": "D" + strconv.Itoa(lastCell), //pvpro
	}

	for _, login := range logins {
		for first, last := range roulette {
			firstCellValue, _ := exFile.GetCellValue(login, first)
			lastCellValue, _ := exFile.GetCellValue(login, last)

			switch first {
			case "B2":
				fmt.Printf("%s wtfskins: $%.2f\n", login, getIncome(firstCellValue, lastCellValue, NOT_PVPRO))
			case "C2":
				fmt.Printf("%s csgolive: $%.2f\n", login, getIncome(firstCellValue, lastCellValue, NOT_PVPRO))
			case "D2":
				fmt.Printf("%s pvpro:    $%.2f (%.0f coins) \n", login,
					getIncome(firstCellValue, lastCellValue, IS_PVPRO),
					getIncome(firstCellValue, lastCellValue, IS_PVPRO)*100)
			}

		}
		fmt.Println()
	}

	fmt.Scanln()
}

func getLastCell(exFile *excelize.File) (int, error) {

	lastCell := 0
	for i := 2; i <= 33; i++ {
		cell, err := exFile.GetCellValue("lemon3100_", "B"+strconv.Itoa(i))

		if err != nil {
			return 0, err
		}

		if cell == "" {
			lastCell = i - 1
		}
	}

	return lastCell, nil
}

func getIncome(first, last string, isPVPRO bool) float64 {

	if !isPVPRO {
		f, _ := strconv.ParseFloat(first[1:], 64)
		l, _ := strconv.ParseFloat(last[1:], 64)

		return l - f
	} else {
		//4433 coins + 560 crystals
		fcoins, _ := strconv.Atoi(first[:strings.Index(first, " ")])
		lcoins, _ := strconv.Atoi(last[:strings.Index(last, " ")])

		return float64(lcoins-fcoins) / 100.0

		/*fcrystals, _ := strconv.Atoi(first[strings.Index(first, "+ ")+2:strings.Index(first, " cr")])
		lcrystals, _ := strconv.Atoi(last[strings.Index(first, "+ ")+2:strings.Index(last, " cr")])

		crystals := lcrystals - fcrystals

		if crystals == 0 {
			return strconv.Itoa(coins)
		} else {
			return strconv.Itoa(coins) + " + " + strconv.Itoa(crystals)
		}*/
	}

}
