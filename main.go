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

var logins = [...]string{"lemon3100_", "rabascal", "destoki2019", "destoki2001", "destoki2002"}

const (
	WTFSKINS_L = "B"
	CSGOLIVE_L = "C"
	PVPRO_L    = "D"
)

type Accounts struct {
	login            string
	wtfskinsFirstVal float64 //dollars
	wtfskinsLastVal  float64
	csgoliveFirstVal float64
	csgoliveLastVal  float64
	pvproFirstVal    int //coins
	pvproLastVal     int
}

func main() {
	exFile, err := excelize.OpenFile("d:\\list.xlsx")

	if err != nil {
		fmt.Println(err)
		fmt.Scanln()
		return
	}

	accounts := []Accounts{
		{login: "l...."},
		{login: "r...."},
		{login: "d...9"},
		{login: "d....1"},
		{login: "de.."},
	}

	checkFirstCells(exFile)
	getLastCellName(exFile, &accounts)

	/*
			IS_PVPRO, NOT_PVPRO := true, false
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

		fmt.Scanln()*/
}

func getLastCellName(exFile *excelize.File, accounts *[]Accounts) {

	getLastCellValue := func(account *Accounts) {

		for i := 'B'; i <= 'D'; i++ {

			for j := 34; j > 2; j-- {
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

	}

	//cuz value is copied but not &
	for i, _ := range *accounts {
		getLastCellValue(&(*accounts)[i])
	}
}

func checkFirstCells(exFile *excelize.File) {

	var buff strings.Builder

	for _, login := range logins {

		var empty bool
		var str string
		cell, _ := exFile.GetCellValue(login, WTFSKINS_L+"2")

		if cell == "" {
			empty = true
			str = "wtfskins "
		}

		cell, _ = exFile.GetCellValue(login, CSGOLIVE_L+"2")

		if cell == "" {
			empty = true
			str += "csgolives "
		}

		cell, _ = exFile.GetCellValue(login, PVPRO_L+"2")

		if cell == "" {
			empty = true
			str += "pvpro"
		}

		if empty {
			buff.WriteString(login + ": " + str + "\n")
		}
	}

	if buff.String() != "" {
		fmt.Printf("Error! List of cells that doesn't contain last month data in the first raw:\n%s", buff.String())
		fmt.Scanln()
		os.Exit(1)
	}
}

/*func getLastCellS(exFile *excelize.File, login, letter string) string {

	lastCell := 0
	for i := 2; i <= 34; i++ {
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

}*/
