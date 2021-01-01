package main

import (
	"log"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type Inventory struct {
	Name      string
	Date      string
	Cake      int
	Chocolate int
	Chips     int
}

func main() {
	inventory := []Inventory{
		{
			Name:      "Bakery1",
			Date:      "26/12/2020",
			Cake:      24,
			Chips:     30,
			Chocolate: 10,
		},
		{
			Name:      "Bakery2",
			Date:      "27/12/2020",
			Cake:      20,
			Chips:     3,
			Chocolate: 10,
		},
		{
			Name:      "Bakery3",
			Date:      "28/12/2020",
			Cake:      2,
			Chips:     3,
			Chocolate: 1,
		},
		{
			Name:      "Bakery4",
			Date:      "29/12/2020",
			Cake:      10,
			Chips:     10,
			Chocolate: 10,
		},
		{
			Name:      "Bakery5",
			Date:      "1/01/2021",
			Cake:      24,
			Chips:     30,
			Chocolate: 10,
		},
	}

	err := SaveToExcel(inventory)
	if err != nil {
		log.Panicf("Error while saving data to excel file: %v", err)
	}

	return
}

func SaveToExcel(inventory []Inventory) error {
	f := excelize.NewFile()

	f.SetCellValue("Sheet1", "A1", "Shop Name")
	f.SetCellValue("Sheet1", "B1", "Date")
	f.SetCellValue("Sheet1", "C1", "Cake")
	f.SetCellValue("Sheet1", "D1", "Chocolate")
	f.SetCellValue("Sheet1", "E1", "Chips")

	for i := 0; i < len(inventory); i++ {
		f.SetCellValue("Sheet1", "A"+strconv.Itoa(i+2), inventory[i].Name)
		f.SetCellValue("Sheet1", "B"+strconv.Itoa(i+2), inventory[i].Date)
		f.SetCellValue("Sheet1", "C"+strconv.Itoa(i+2), inventory[i].Cake)
		f.SetCellValue("Sheet1", "D"+strconv.Itoa(i+2), inventory[i].Chocolate)
		f.SetCellValue("Sheet1", "E"+strconv.Itoa(i+2), inventory[i].Chips)

	}

	if err := f.SaveAs("inventory.xlsx"); err != nil {
		return err
	}

	return nil
}
