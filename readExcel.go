package main

import (
	"fmt"
	"os"
	"strconv"
	"strings"
	"sync"

	"github.com/fatih/color"
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

//Check if file or directory exists
func fileExists(path string) bool {
	if _, err := os.Stat(path); err == nil {
		return true
	}

	return false
}

//Get path to the directory that exactly doesn't exist in the whole path
func nonExistedFirstDir(path string) (string, error) {
	//"D:\\Program Files\\MEGAsync\\MEGAsync\\Internet Deals\\Steam\\ルーレット"

	paths := []string{path} //1st element is original path and in the loop we reduce whole path by 1 directory

	for {
		if !strings.Contains(path, "\\") {
			break
		}

		path = path[:strings.LastIndex(path, "\\")]
		paths = append(paths, path)
	}

	//D:\\, D:\\Program Files, ... -> the last element is the 1st to check
	for i := len(paths) - 1; i >= 0; i-- {
		if !fileExists(paths[i]) {
			return paths[i], nil
		}
	}

	return "", fmt.Errorf("error in nonExistedFirstDir(): all directories along %s exist", path)
}
