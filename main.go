package main

import (
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

/*
	1. Начального значения если нет - ошибка  == ибо это первый месяц
	2. Проверить, чтобы у всех клеточек было проверено их максимальное и начальное значение  - отнять только от последнего значения,
		если его нет  - то значит месяц пустой  = (если хотя бы у одного аккаунта есть данные значит это не ошибка, а $0.00)
	3. Структура
*/

func main() {
	exFile, err := excelize.OpenFile("d:\\list.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	IS_PVPRO, NOT_PVPRO := true, false
	logins := [...]string{"l...", "r...l", "d...9", "..001", "..."}

	roulette := map[string]string{
		"B2": "B", //wtfskins
		"C2": "C", //csgolines
		"D2": "D", //pvpro
	}

	for _, login := range logins {
		for firstCellName, letter := range roulette {
			firstCellValue, _ := exFile.GetCellValue(login, firstCellName)
			lastCellValue, _ := exFile.GetCellValue(login, getLastCellS(exFile, login, letter))

			switch firstCellName {
			case "B2":
				fmt.Printf("%s wtfskins: $%.2f\n", login,
					getIncome(exFile, login, letter, firstCellValue, lastCellValue, NOT_PVPRO))
			case "C2":
				fmt.Printf("%s csgolive: $%.2f\n", login,
					getIncome(exFile, login, firstCellName, letter, firstCellValue, lastCellValue, NOT_PVPRO))
			case "D2":
				fmt.Printf("%s pvpro:    $%.2f (%.0f coins) \n", login,
					getIncome(exFile, login, letter, firstCellValue, lastCellValue, IS_PVPRO),
					getIncome(exFile, login, letter, firstCellValue, lastCellValue, IS_PVPRO)*100)
			}

		}
		fmt.Println()
	}

	fmt.Scanln()
}

func getLastCellS(exFile *excelize.File, login, letter string) string {

	lastCell := 0
	for i := 2; i <= 33; i++ {
		cell, err := exFile.GetCellValue(login, letter+strconv.Itoa(i))

		if err != nil {
			fmt.Println(err)
			os.Exit(1)
		}

		if cell == "" {
			lastCell = i - 1
		}
	}

	return letter + strconv.Itoa(lastCell)
}

func getIncome(exFile *excelize.File, login, letter, firstVal, lastVal string, isPVPRO bool) float64 {

	if !isPVPRO {

		i, _ := strconv.Atoi(first[1:])
		n, _ := strconv.Atoi(last[1:])

		var additional float64 = 0

		for ; i <= n; i++ {
			cell, _ := exFile.GetCellValue(login, letter+strconv.Itoa(i))

			if strings.Contains(cell, "->") {
				//$0.29 -> $0.02
				part1, _ := strconv.ParseFloat(cell[1:strings.Index(cell, " ->")], 64)
				part2, _ := strconv.ParseFloat(cell[strings.Index(cell, " $")+2:], 64)

				additional += (part1 - part2)
			}

		}

		f, _ := strconv.ParseFloat(firstVal[1:], 64)
		l, _ := strconv.ParseFloat(lastVal[1:], 64)

		return l - f + additional
	} else {
		//4433 coins + 560 crystals
		fcoins, _ := strconv.Atoi(first[:strings.Index(first, " ")])
		lcoins, _ := strconv.Atoi(last[:strings.Index(last, " ")])

		return float64(lcoins-fcoins) / 100.0
	}

}
