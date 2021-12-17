package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"os"
	"runtime"
	"strconv"
	"strings"
	"time"

	"github.com/fatih/color"
	"github.com/xuri/excelize/v2"
)

//Key function for working with 2nd file -> _ルーレットの総収入.xlsx
func writeToFileOverallIncome(workFolder, fileTotalIncomeName, fileMonthName string, totalOverallIncomeInDollars float64,
	incomeMonth, incomeYear int, fileMonthNameAlias string) {

	fmt.Println()

	defer func() {
		fmt.Println(strings.Repeat("_", 110))
	}()

	var (
		exFileIncome *excelize.File
		err          error
		data         []byte
		date         string
	)

	if exFileIncome, err = getExcelFileIncome(workFolder, fileTotalIncomeName); err != nil {
		fmt.Println(err)
		return
	}

	sheetsNumber := 0
	for {
		incomeSheetName := exFileIncome.GetSheetName(sheetsNumber)
		if incomeSheetName == "" {
			break
		}
		sheetsNumber++
	}
	sheetYearIndex := incomeYear - 2020 //2020 - first sheet (at index 0)

	if data, err = getPrivatAPIData(incomeMonth, incomeYear, &date); err != nil {
		fmt.Println(err)
		return
	}

	var hryvniaRate, rubleRate, totalOverallIncomeInHryvnias, totalOverallIncomeInRubles float64
	if err = getCurrencyRate(data, &hryvniaRate, &rubleRate, &totalOverallIncomeInHryvnias,
		&totalOverallIncomeInRubles, &totalOverallIncomeInDollars); err != nil {

		fmt.Println(err)
		return
	}

	//stores old and new values of cells
	var beforeS, afterS []string

	if err = setMonthCellsInOverallIncomeFile(exFileIncome, sheetYearIndex, incomeMonth, totalOverallIncomeInDollars, totalOverallIncomeInRubles,
		totalOverallIncomeInHryvnias, &beforeS, &afterS); err != nil {

		fmt.Println(err)
		return
	}

	var yearIncomeInDollars, yearIncomeInRubles, yearIncomeInHryvnias float64
	calculateAnnualIncome(exFileIncome, sheetYearIndex, &yearIncomeInDollars, &yearIncomeInRubles, &yearIncomeInHryvnias)

	if len(beforeS) > 0 {
		color.Red(strings.Join(beforeS, "\t"))
		color.Cyan(strings.Join(afterS, "\t"))
	}

	//Reset values for reuse
	beforeS = nil
	afterS = nil

	if err = setYearCellsInOverallIncomeFile(exFileIncome, sheetYearIndex, yearIncomeInDollars,
		yearIncomeInRubles, yearIncomeInHryvnias, &beforeS, &afterS); err != nil {

		fmt.Println(err)
		return
	}

	if len(beforeS) > 0 {
		color.Yellow("Income ($, ₽, ₴):  $1 = ₴%.2f = ₽%.2f [НБУ, %s]", hryvniaRate, hryvniaRate/rubleRate, date)
		fmt.Print("\n\n")
		color.Yellow("Annual income ($, ₽, ₴):")

		color.Red("\t" + strings.Join(beforeS, "  \t"))
		color.Cyan("\t" + strings.Join(afterS, "  \t"))

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

		if err := exFileIncome.Save(); err != nil {
			color.Red("\n%s", err)
			color.Red("\n\nCalculated values have not been stored into %s", fileMonthNameAlias)
		} else {
			color.Green("\n\nCalculated values have been successfully stored into %s", fileMonthNameAlias)
		}
	}
}

//choose only one .xlsx file to work with from of two directories
func getExcelFileIncome(workFolder, fileTotalIncomeName string) (*excelize.File, error) {
	//search files both in workFolder and currFolder
	exFileIncome, err := excelize.OpenFile(workFolder + "\\" + fileTotalIncomeName)

	if err != nil {
		var err2 error

		currFolder, _ := os.Getwd()
		if runtime.GOOS == "windows" {
			exFileIncome, err2 = excelize.OpenFile(currFolder + "\\" + fileTotalIncomeName)
		} else {
			exFileIncome, err2 = excelize.OpenFile(currFolder + "/" + fileTotalIncomeName)
		}

		if err2 != nil {
			return nil, fmt.Errorf("%s\n%s\n", err, err2)
		}
	}

	return exFileIncome, nil
}

//get raw request from Privat24 CurrencyExchange API
func getPrivatAPIData(incomeMonth, incomeYear int, date *string) ([]byte, error) {

	//if we check roulettes data of previous month we choose 28th day, if current month - so current day
	if time.Now().Month() == time.Month(incomeMonth) && time.Now().Day() <= 28 {
		*date = time.Now().Format("02.01.2006")
	} else {
		//28.11.2021, 28.02.2022, 28.05.2022 -> cuz 28th day exists in each month
		*date = "28." + fmt.Sprintf("%02d", incomeMonth) + "." + strconv.Itoa(incomeYear)
	}

	resp, err := http.Get("https://api.privatbank.ua/p24api/exchange_rates?json&date=" + *date)

	if err != nil {
		return nil, err
	}

	defer resp.Body.Close()

	data, _ := ioutil.ReadAll(resp.Body)

	//remove useless from the beginning ""date":"09.12.2021","bank":"PB","baseCurrency":980,"baseCurrencyLit":"UAH","exchangeRate":["
	//privatData := []byte(string(buff)[strings.Index(string(buff), "[")+1:])
	data = []byte(string(data)[strings.Index(string(data), "[") : len(data)-1])

	return data, nil
}

//get currency rate from raw API data through json and struct
func getCurrencyRate(data []byte, hryvniaRate, rubleRate, totalOverallIncomeInHryvnias,
	totalOverallIncomeInRubles, totalOverallIncomeInDollars *float64) error {

	//json line: "baseCurrency":"UAH","currency":"RUB","saleRateNB":0.3692500,"purchaseRateNB":0.3692500,"saleRate":0.3860000,"purchaseRate":0.3560000
	type Currency struct {
		ForeignCurrencyName string  `json:"currency"`
		PurchaseRateNBU     float64 `json:"purchaseRateNB"`
	}

	currencies := []Currency{}
	if err := json.Unmarshal(data, &currencies); err != nil {
		return err
	}

	for _, currency := range currencies {
		if currency.ForeignCurrencyName == "USD" {
			*hryvniaRate = currency.PurchaseRateNBU //how many hryvnias in $1
			*totalOverallIncomeInHryvnias = (*hryvniaRate) * (*totalOverallIncomeInDollars)
		}
	}

	//2 loops cuz we need to get income in hryvnia first
	for _, currency := range currencies {
		if currency.ForeignCurrencyName == "RUB" {
			*rubleRate = currency.PurchaseRateNBU
			*totalOverallIncomeInRubles = (*totalOverallIncomeInHryvnias) / (*rubleRate)
		}
	}

	return nil
}

//set new value new values to the current month and if old and new values are different -> information is stored to beforeS, afterS
func setMonthCellsInOverallIncomeFile(exFileIncome *excelize.File, sheetYearIndex, incomeMonth int,
	totalOverallIncomeInDollars, totalOverallIncomeInRubles, totalOverallIncomeInHryvnias float64, beforeS, afterS *[]string) error {
	//Set value for the Dollars cell
	if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "B"+strconv.Itoa(incomeMonth+1)); err != nil {
		fmt.Println(err)
	} else {
		totalOverallIncomeInDollarsS := fmt.Sprintf("$%.2f", totalOverallIncomeInDollars)
		if val != totalOverallIncomeInDollarsS {
			if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "B"+strconv.Itoa(incomeMonth+1),
				totalOverallIncomeInDollarsS); err != nil {
				return err
			} else {
				*beforeS = append(*beforeS, "\t"+val)
				*afterS = append(*afterS, "\t"+totalOverallIncomeInDollarsS)
			}
		}
	}

	//Set value for the Rubles cell
	if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "C"+strconv.Itoa(incomeMonth+1)); err != nil {
		fmt.Println(err)
	} else {
		totalOverallIncomeInRublesS := fmt.Sprintf("₽%.2f", totalOverallIncomeInRubles)
		if val != totalOverallIncomeInRublesS {
			if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "C"+strconv.Itoa(incomeMonth+1),
				totalOverallIncomeInRublesS); err != nil {
				return err
			} else {
				*beforeS = append(*beforeS, "\t"+val)
				*afterS = append(*afterS, "\t"+totalOverallIncomeInRublesS)
			}
		}
	}

	//Set value for the Hryvnia cell
	if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "D"+strconv.Itoa(incomeMonth+1)); err != nil {
		fmt.Println(err)
	} else {
		totalOverallIncomeInHryvniasS := fmt.Sprintf("₴%.2f", totalOverallIncomeInHryvnias)
		if val != totalOverallIncomeInHryvniasS {
			if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "D"+strconv.Itoa(incomeMonth+1),
				totalOverallIncomeInHryvniasS); err != nil {
				return err
			} else {
				*beforeS = append(*beforeS, "\t"+val)
				*afterS = append(*afterS, "\t"+totalOverallIncomeInHryvniasS)
			}
		}
	}

	return nil
}

//calculate income of each month in dollars, rubles and hryvnias
func calculateAnnualIncome(exFileIncome *excelize.File, sheetYearIndex int, yearIncomeInDollars, yearIncomeInRubles, yearIncomeInHryvnias *float64) {
	for i := 'B'; i <= 'D'; i++ {
		for j := 2; j <= 13; j++ {
			if i == 'B' {
				if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "B"+strconv.Itoa(j)); err != nil {
					fmt.Println(err)
				} else {
					if len(val) > 0 { // in case of an empty string
						if num, err := strconv.ParseFloat(val[1:], 64); err == nil { //$ = 1 byte ($11.80)
							*yearIncomeInDollars += num
						}
					}
				}
			} else if i == 'C' {
				if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "C"+strconv.Itoa(j)); err != nil {
					fmt.Println(err)
				} else {
					if len(val) > 0 {
						if num, err := strconv.ParseFloat(val[3:], 64); err == nil { //₽ = 3 bytes (₽881.92)
							*yearIncomeInRubles += num
						}
					}
				}
			} else if i == 'D' {
				if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "D"+strconv.Itoa(j)); err != nil {
					fmt.Println(err)
				} else {
					if len(val) > 0 {
						if num, err := strconv.ParseFloat(val[3:], 64); err == nil { //₴ = 3 bytes (₴318.90)
							*yearIncomeInHryvnias += num
						}
					}
				}
			}
		}
	}
}

//set new values to '全ての月' and if old and new values are different -> information is stored to beforeS, afterS
func setYearCellsInOverallIncomeFile(exFileIncome *excelize.File, sheetYearIndex int,
	yearIncomeInDollars, yearIncomeInRubles, yearIncomeInHryvnias float64, beforeS, afterS *[]string) error {
	//Set value for the YearDollars cell
	if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "B14"); err != nil {
		fmt.Println(err)
	} else {
		yearIncomeInDollarsS := fmt.Sprintf("$%.2f", yearIncomeInDollars)
		if val != yearIncomeInDollarsS {
			if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "B14", yearIncomeInDollarsS); err != nil {
				return err
			} else {
				*beforeS = append(*beforeS, val)
				*afterS = append(*afterS, yearIncomeInDollarsS)
			}
		}
	}

	//Set value for the YearRubles cell
	if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "C14"); err != nil {
		fmt.Println(err)
	} else {
		yearIncomeInRublesS := fmt.Sprintf("₽%.2f", yearIncomeInRubles)
		if val != yearIncomeInRublesS {
			if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "C14", yearIncomeInRublesS); err != nil {
				return err
			} else {
				*beforeS = append(*beforeS, val)
				*afterS = append(*afterS, yearIncomeInRublesS)
			}
		}
	}

	//Set value for the YearHryvnia cell
	if val, err := exFileIncome.GetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "D14"); err != nil {
		fmt.Println(err)
	} else {
		yearIncomeInHryvniasS := fmt.Sprintf("₴%.2f", yearIncomeInHryvnias)
		if val != yearIncomeInHryvniasS {
			if err = exFileIncome.SetCellValue(exFileIncome.GetSheetName(sheetYearIndex), "D14", yearIncomeInHryvniasS); err != nil {
				return err
			} else {
				*beforeS = append(*beforeS, val)
				*afterS = append(*afterS, yearIncomeInHryvniasS)
			}
		}
	}

	return nil
}
