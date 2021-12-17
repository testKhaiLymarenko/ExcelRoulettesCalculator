package main

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/fatih/color"
	"github.com/xuri/excelize/v2"
)

//Key function for work with 1st file -> (Example: 2021年12月のルーレット.xlsx)
func writeToFileMonthIncome(exFile *excelize.File, accounts *[]Accounts, totalWtfskinsIncome, totalCsgolivesIncome,
	totalPvproDollarsIncome, totalOverallIncomeInDollars float64, totalPvproCoinsIncome int,
	fileMonthName, fileMonthNameAlias string) {

	defer func() {
		fmt.Println(strings.Repeat("_", 110))
	}()

	incomeSheetName := exFile.GetSheetName(0)
	var beforeS, afterS []string //store current cell values in beforeS slice, and new values in afterS

	if err := setRoulettesCellsInMonthIncomeFile(exFile, accounts, incomeSheetName, &beforeS, &afterS); err != nil {
		fmt.Println(err)
		return
	}

	//if current and new values are different
	if len(beforeS) > 0 {
		printDifferenceBetweenCellValues(accounts, &beforeS, &afterS)
	}

	//reset slices to use it for OVERALL cells
	beforeS = nil
	afterS = nil

	if err := setCurrentMonthCellValues(exFile, incomeSheetName, &beforeS, &afterS, totalWtfskinsIncome, totalCsgolivesIncome,
		totalPvproDollarsIncome, totalOverallIncomeInDollars, totalPvproCoinsIncome); err != nil {
		fmt.Println(err)
		return
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
}

//write data to account cells: wtfskins, csgolive, pvpro
func setRoulettesCellsInMonthIncomeFile(exFile *excelize.File, accounts *[]Accounts, incomeSheetName string,
	beforeS, afterS *[]string) error {

	for i := 'B'; i <= 'D'; i++ {
		for j := range *accounts {

			if i == 'B' {
				//Check if current and storead values are the same, if they are different - overwrite it here (it will apply)
				//	only if we call Save() and store these values in 2 different slices to show the difference later
				if val, err := exFile.GetCellValue(incomeSheetName, "B"+strconv.Itoa(j+2)); err != nil {
					return err
				} else {
					accountWtfskinsIncomeS := fmt.Sprintf("+$%.2f", (*accounts)[j].wtfskinsIncome)

					if val != accountWtfskinsIncomeS {
						if err := exFile.SetCellValue(incomeSheetName, "B"+strconv.Itoa(j+2), accountWtfskinsIncomeS); err != nil {
							return err
						} else {
							*beforeS = append(*beforeS, "wtfskins: "+val)
							*afterS = append(*afterS, "wtfskins: "+accountWtfskinsIncomeS)
						}
					}

				}
			} else if i == 'C' {
				if val, err := exFile.GetCellValue(incomeSheetName, "C"+strconv.Itoa(j+2)); err != nil {
					return err
				} else {
					accountCsgoliveIncomeS := fmt.Sprintf("+$%.2f", (*accounts)[j].csgoliveIncome)

					if val != accountCsgoliveIncomeS {
						if err := exFile.SetCellValue(incomeSheetName, "C"+strconv.Itoa(j+2), accountCsgoliveIncomeS); err != nil {
							return err
						} else {
							*beforeS = append(*beforeS, "csgolive: "+val)
							*afterS = append(*afterS, "csgolive: "+accountCsgoliveIncomeS)
						}
					}

				}
			} else if i == 'D' {
				if val, err := exFile.GetCellValue(incomeSheetName, "D"+strconv.Itoa(j+2)); err != nil {
					return err
				} else {
					accountPvproIncomeCoinsS := fmt.Sprintf("+%d coins", (*accounts)[j].pvproCoinsIncome)

					if val != accountPvproIncomeCoinsS {
						if err := exFile.SetCellValue(incomeSheetName, "D"+strconv.Itoa(j+2), accountPvproIncomeCoinsS); err != nil {
							return err
						} else {
							*beforeS = append(*beforeS, "pvpro: "+val)
							*afterS = append(*afterS, "pvpro: "+accountPvproIncomeCoinsS)
						}
					}

				}
			}
		}
	}

	return nil
}

func printDifferenceBetweenCellValues(accounts *[]Accounts, beforeS, afterS *[]string) {
	fmt.Print("\n\n")

	//account1: [0], [1], [3] | account2: [4], [5], [6] | ....
	beforeSlice := make([][]string, len(*accounts))
	afterSlice := make([][]string, len(*accounts))
	//index for loop in accounts names
	var j int

	for i := 0; i < len(*beforeS); i++ {

		//wtfskins, csgolive and pvpro values changes every N elements
		if i%len(*accounts) == 0 {
			j = 0
		}

		beforeSlice[j] = append(beforeSlice[j], (*beforeS)[i])
		afterSlice[j] = append(afterSlice[j], (*afterS)[i])

		j++
	}

	//print values in just one line
	for i := 0; i < len(*accounts); i++ {
		color.Yellow((*accounts)[i].login + ": ")
		color.Red("\t" + strings.Join(beforeSlice[i], "\t"))
		color.Cyan("\t" + strings.Join(afterSlice[i], "\t"))
	}
}

//check if current and storead values are the same, if they are different - overwrite it here (it will apply)
//only if we call Save() and store these values in 2 different slices to show the difference later
func setCurrentMonthCellValues(exFile *excelize.File, incomeSheetName string, beforeS, afterS *[]string,
	totalWtfskinsIncome, totalCsgolivesIncome, totalPvproDollarsIncome, totalOverallIncomeInDollars float64, totalPvproCoinsIncome int) error {
	if val, err := exFile.GetCellValue(incomeSheetName, "B8"); err != nil {
		return err
	} else {
		totalWtfskinsIncomeS := fmt.Sprintf("+$%.2f", totalWtfskinsIncome)
		if val != totalWtfskinsIncomeS {
			if err := exFile.SetCellValue(incomeSheetName, "B8", totalWtfskinsIncomeS); err != nil {
				return err
			} else {
				*beforeS = append(*beforeS, "\twtfskins: "+val)
				*afterS = append(*afterS, "\twtfskins: "+totalWtfskinsIncomeS)
			}
		}
	}

	if val, err := exFile.GetCellValue(incomeSheetName, "C8"); err != nil {
		return err
	} else {
		totalCsgolivesIncomeS := fmt.Sprintf("+$%.2f", totalCsgolivesIncome)
		if val != totalCsgolivesIncomeS {
			if err := exFile.SetCellValue(incomeSheetName, "C8", totalCsgolivesIncomeS); err != nil {
				return err
			} else {
				*beforeS = append(*beforeS, "\tcsgolive: "+val)
				*afterS = append(*afterS, "\tcsgolive: "+totalCsgolivesIncomeS)
			}
		}
	}

	if val, err := exFile.GetCellValue(incomeSheetName, "D8"); err != nil {
		return err
	} else {
		totalPvproIncomeS := fmt.Sprintf("+%d coins (+$%.2f)", totalPvproCoinsIncome, totalPvproDollarsIncome)
		if val != totalPvproIncomeS {
			if err := exFile.SetCellValue(incomeSheetName, "D8", totalPvproIncomeS); err != nil {
				return err
			} else {
				*beforeS = append(*beforeS, "\tpvpro: "+val)
				*afterS = append(*afterS, "\tpvpro: "+totalPvproIncomeS)
			}
		}
	}

	if val, err := exFile.GetCellValue(incomeSheetName, "C11"); err != nil {
		return err
	} else {
		totalOverallIncomeInDollarsS := fmt.Sprintf("Total Income: $%.2f", totalOverallIncomeInDollars)
		if val != totalOverallIncomeInDollarsS {
			if err := exFile.SetCellValue(incomeSheetName, "C11", totalOverallIncomeInDollarsS); err != nil {
				return err
			} else {
				*beforeS = append(*beforeS, "\t"+val)
				*afterS = append(*afterS, "\t"+totalOverallIncomeInDollarsS)
			}
		}
	}

	return nil
}
