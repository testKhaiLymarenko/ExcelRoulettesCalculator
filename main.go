package main

import (
	"fmt"
	"os"
	"strconv"
	"strings"
	"sync"

	"github.com/xuri/excelize/v2"
)

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

	a.wtfskinsIncome = a.wtfskinsLastVal - a.wtfskinsFirstVal
	a.csgoliveIncome = a.csgoliveLastVal - a.csgoliveFirstVal
	a.pvproCoinsIncome = a.pvproLastVal - a.pvproFirstVal
	a.pvproDollarsIncome = float64(a.pvproCoinsIncome) / 1000

	return "\t\t\t" + a.login + ":\nwtfskins: " + fmt.Sprintf("$%.2f\n", a.wtfskinsIncome) + "csgolive: " +
		fmt.Sprintf("$%.2f\n", a.csgoliveIncome) + "pvpro:    $" + fmt.Sprintf("%.2f (%d coins)\n", a.pvproDollarsIncome, a.pvproCoinsIncome)
}

//сделать: ошибка, когда нет первого значения: wtfskins: d..., d..., l.., и т.д

func main() {
	exFile, err := excelize.OpenFile("d:\\list.xlsx")

	if err != nil {
		fmt.Println(err)
		fmt.Scanln()
		return
	}

	accounts := []Accounts{
		{login: "....."},
		{login: "...."},
		{login: "d...."},
		{login: "...."},
		{login: "d....2"},
	}

	var totalWtfskinsIncome, totalCsgolivesIncome, totalPvproDollarsIncome float64
	var totalPvproCoinsIncome int

	checkAndGetFirstCells(exFile, &accounts)
	getLastCellName(exFile, &accounts)

	for _, account := range accounts {
		fmt.Println(account.CalculateS())

		totalWtfskinsIncome += account.wtfskinsIncome
		totalCsgolivesIncome += account.csgoliveIncome

		totalPvproCoinsIncome += account.pvproCoinsIncome
		totalPvproDollarsIncome += account.pvproDollarsIncome
	}

	fmt.Printf("\n\t\t\tTotal Income (%d accounts):\nwtfskins:  $%.2f\ncsgolives: $%.2f\npvpro:     $%.2f (%d coins)\n\nOverall:   $%.2f\n",
		len(accounts), totalWtfskinsIncome, totalCsgolivesIncome, totalPvproDollarsIncome, totalPvproCoinsIncome,
		(totalWtfskinsIncome + totalCsgolivesIncome + totalPvproDollarsIncome))

	fmt.Scanln()
}

func getLastCellName(exFile *excelize.File, accounts *[]Accounts) {

	var wg sync.WaitGroup

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

		wg.Done()
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

	var buff strings.Builder

	//B - wtfskins, C - csgolive, D - pvpro
	for i := 'B'; i <= 'D'; i++ {
		//check only one column of each account
		for j := 0; j < len(*accounts); j++ {
			login := (*accounts)[j].login
			//need to
			empty := false

			cell, _ := exFile.GetCellValue(login, string(i)+"2") //B + 2 = B2 cell

			if i == 'B' {
				if cell == "" { //
					empty = true
					buff.WriteString(login + ": wtfskins")
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
					buff.WriteString(login + ": csgolives")
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
					buff.WriteString(login + ": pvpro")
				} else {
					if val, err := strconv.Atoi(cell[:strings.Index(cell, " ")]); err == nil {
						(*accounts)[j].pvproFirstVal = val
					} else {
						fmt.Println(err)
					}
				}
			}

			if empty {
				buff.WriteString("\n")
			}
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
