/*
Copyright © 2021 NAME HERE <EMAIL ADDRESS>

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/
package cmd

import (
	"encoding/csv"
	"errors"
	"fmt"
	"io"
	"log"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/spf13/viper"
	"github.com/xuri/excelize/v2"
)

// Flags
var account string
var salesChannel string
var serviceCode string
var originCountry string
var incoterm string
var description string
var HSCode string
var unitWeightStr string
var weightMin string
var weightMax string

var EXPORT_HEADERS = []string{
	"Pick-up Account Number",
	"Sales Channel",
	"Shipment Order ID",
	"Tracking Number",
	"Shipping Service Code",
	"Company",
	"Consignee Name",
	"Address Line 1",
	"Address Line 2",
	"Address Line 3",
	"City",
	"State",
	"Postal Code",
	"Destination Country Code",
	"Phone Number",
	"Email Address",
	"Consignee Fiscal ID",
	"Consignee Fiscal ID Type",
	"Consignee ID for Alternative Delivery Location",
	"Shipment Weight (g)",
	"Length (cm)",
	"Width (cm)",
	"Height (cm)",
	"Currency Code",
	"Total Declared Value",
	"incoterm",
	"Tax Prepaid",
	"Order URL",
	"E-Seller Platform Name",
	"Freight",
	"Is Insured",
	"Insurance",
	"Is COD",
	"Cash on Delivery Value",
	"Recipient ID",
	"Recipient ID Type",
	"Duties",
	"Taxes",
	"Workshare Indicator",
	"Shipment Description",
	"Shipment Import Description",
	"Shipment Export Description",
	"Shipment Content Indicator",
	"Content Description",
	"Content Import Description",
	"Content Export Description",
	"Content Unit Price",
	"Content Origin Country",
	"Content Quantity",
	"Content Weight (g)",
	"Content Code",
	"HS Code",
	"Content Indicator",
	"Marketplace Name",
	"Marketplace URL",
	"Remarks",
	"Customs Certificate",
	"Customs Licence",
	"Customs Invoice Number",
	"Customs Sales Order Number",
	"Customs Purchase Order Number",
	"Shipper Company",
	"Shipper Name",
	"Shipper Address1",
	"Shipper Address2",
	"Shipper Address3",
	"Shipper City",
	"Shipper State",
	"Shipper Postal Code",
	"Shipper CountryCode",
	"Shipper Phone Number",
	"Shipper Email address",
	"Shipper Fiscal ID",
	"Shipper Fiscal ID Type",
	"GstIN",
	"IEC Number",
	"Bank AD Code",
	"Bank AC",
	"Bank IFSC",
	"Return Shipping Service Code",
	"Return Company",
	"Return Name",
	"Return Address Line 1",
	"Return Address Line 2",
	"Return Address Line 3",
	"Return City",
	"Return State",
	"Return Postal Code",
	"Return Destination Country Code",
	"Return Phone Number",
	"Return Email Address",
	"Return Fiscal ID",
	"Return Fiscal ID Type",
	"Service1",
	"Service2",
	"Service3",
	"Service4",
	"Service5",
	"Grouping Reference1",
	"Grouping Reference2",
	"Customer Reference 1",
	"Customer Reference 2",
	"Handover Method",
	"Return Mode",
	"Billing Reference 1",
	"Billing Reference 2",
	"IsMult",
	"Delivery Option",
	"PieceID",
	"Piece Description",
	"Piece Weight",
	"Piece COD",
	"Piece Insurance",
	"Piece Billing Reference 1",
	"Piece Billing Reference 2",
	"Invoice Number",
	"Invoice Date",
	"CGST Amount",
	"SGST Amount",
	"IGST Amount",
	"CESS Amount",
	"IGST Rate %",
	"MEIS",
	"Commodity Under 3C",
	"Discount",
	"Reverse Charge",
	"IGST Payment Status",
	"Reason For Export",
	"Terms of Invoice",
	"Description of Other Charges",
	"Amount of Other Charges",
}
var EXPORT_HEADERS_MAP = map[string]string{
	"Shipping Name":          "Consignee Name",
	"Shipping Address1":      "Address Line 1",
	"Shipping Address2":      "Address Line 2",
	"Shipping City":          "City",
	"Shipping Province Name": "State",
	"Shipping Country":       "Destination Country Code",
	"Shipping Phone":         "Phone Number",
	"Email":                  "Email Address",
	"Currency":               "Currency Code",
	"Total":                  "Total Declared Value",
	"Taxes":                  "Taxes",
	"Shipping Zip":           "Postal Code",
}

func check(e error) {
	if e != nil {
		panic(e)
	}
}

func copyFile(in, out string) {
	i, e := os.Open(in)
	check(e)
	defer i.Close()
	o, e := os.Create(out)
	check(e)
	defer o.Close()
	_, e = io.Copy(o, i)
	check(e)
	e = o.Sync()
	check(e)
}

func readCsvFile(filePath string) [][]string {
	f, err := os.Open(filePath)
	check(err)
	defer f.Close()

	csvReader := csv.NewReader(f)

	records, err := csvReader.ReadAll()
	check(err)
	return records
}

// convertCmd represents the convert command
var convertCmd = &cobra.Command{
	Use:   "convert [order-export.csv] [output; Default=output.txt]",
	Short: "Convert Shopify orders export to bulk upload template",
	Long: `Convert Shopify orders export to bulk upload template:

Select your orders you want to export on Shopify. Export selected as CSV.`,
	Args: func(cmd *cobra.Command, args []string) error {
		if len(args) < 1 {
			return errors.New("requires an order export argument")
		}
		if len(args) < 2 {
			return errors.New("requires an dhl excel argument")
		}
		if _, err := os.Stat(args[0]); os.IsNotExist(err) {
			return fmt.Errorf("invalid order export specified does not exist: %s", args[0])
		}
		return nil
	},
	Run: func(cmd *cobra.Command, args []string) {
		input := readCsvFile(args[0])
		unitWeight, err := strconv.ParseFloat(unitWeightStr, 64)
		if err != nil {
			log.Fatal("Failed exporting")
		}

		ExportValues := map[string]map[string]string{}
		for _, orderInput := range input[1:] {
			orderId := orderInput[0]
			_, ok := ExportValues[orderId]
			if !ok {
				newOrder := map[string]string{}
				newOrder["Pick-up Account Number"] = account
				newOrder["Sales Channel"] = salesChannel
				newOrder["Shipping Service Code"] = serviceCode
				newOrder["Content Origin Country"] = originCountry
				newOrder["incoterm"] = incoterm
				newOrder["Shipment Description"] = description
				newOrder["HS Code"] = HSCode

				newOrder["Shipment Order ID"] = orderId
				ExportValues[orderId] = newOrder
			}

			order, ok := ExportValues[orderId]
			if !ok {
				log.Fatal("Failed exporting")
			}

			for j, orderValue := range orderInput {
				key := input[0][j]
				if orderValue == "" {
					continue
				}
				if key == "Name" {
					order["Shipment Order ID"] = strings.Replace(orderValue, "#", "", 1)
				} else if key == "Lineitem quantity" {
					quantity, err := strconv.Atoi(orderValue)
					check(err)
					if val, ok := order["Content Quantity"]; ok {
						currentQuantity, err := strconv.Atoi(val)
						check(err)
						order["Content Quantity"] = fmt.Sprintf("%d", quantity+currentQuantity)
					} else {
						order["Content Quantity"] = fmt.Sprintf("%d", quantity)
					}

					itemWeight := quantity * int(unitWeight)
					if val, ok := order["Shipment Weight (g)"]; ok {
						itemTotalWeight, err := strconv.Atoi(val)
						check(err)
						order["Shipment Weight (g)"] = fmt.Sprintf("%d", itemWeight+itemTotalWeight)
					} else {
						order["Shipment Weight (g)"] = fmt.Sprintf("%d", itemWeight)
					}
				} else if key == "Lineitem sku" {
					// TODO: Handle multiple line items
					if _, ok := order["Content Code"]; !ok {
						order["Content Code"] = orderValue
					}
				} else if key == "Lineitem name" {
					// TODO: Handle multiple line items
					if _, ok := order["Content Description"]; !ok {
						if len(orderValue) > 49 {
							order["Content Description"] = orderValue[:49]
						} else {
							order["Content Description"] = orderValue
						}
					}
				} else if val, ok := EXPORT_HEADERS_MAP[key]; ok {
					// fmt.Println("mapped", key, ":", val, "=", orderValue)
					if key == "Shipping Zip" || key == "Shipping Phone" {
						order[val] = strings.Replace(orderValue, "'", "", 1)
					} else {
						order[val] = orderValue
					}
				}
			}
			ExportValues[orderId] = order
		}

		fmt.Println("\nExported", len(ExportValues), "orders")

		now := time.Now()
		outputFile := fmt.Sprintf("output-%d.xlsx", now.UnixNano())
		copyFile(args[1], outputFile)
		excel, err := excelize.OpenFile(outputFile)
		check(err)

		// Write values
		i := 2
		for orderId, order := range ExportValues {
			// row
			for col, field := range EXPORT_HEADERS {
				cell, _ := excelize.CoordinatesToCellName(col+1, i)
				if field == "Content Unit Price" {
					qty, err := strconv.Atoi(order["Content Quantity"])
					check(err)
					weight, err := strconv.Atoi(order["Shipment Weight (g)"])
					check(err)

					unitWeight := weight / qty
					weightWarn := false
					if weightMin != "" {
						weightThreshold, err := strconv.Atoi(weightMin)
						check(err)

						if unitWeight < weightThreshold {
							weightWarn = true
							fmt.Printf("[%s] Possible under weight, %dg avg unit weight\n", orderId, unitWeight)
						}
					}
					if !weightWarn && weightMax != "" {
						weightThreshold, err := strconv.Atoi(weightMax)
						check(err)

						if unitWeight > weightThreshold {
							fmt.Printf("[%s] Possible over weight, %dg avg unit weight\n", orderId, unitWeight)
						}
					}
					totalValue, err := strconv.ParseFloat(order["Total Declared Value"], 64)
					check(err)
					excel.SetCellValue("Sheet1", cell, fmt.Sprintf("%0.2f", totalValue/float64(qty)))
				} else if val, ok := order[field]; ok {
					excel.SetCellValue("Sheet1", cell, val)
				}
			}
			i += 1
		}
		if err = excel.Save(); err != nil {
			fmt.Println(err)
		}
	},
}

func init() {
	rootCmd.PersistentFlags().StringVar(&account, "account", "", "DHL Pickup account number")
	viper.BindPFlag("account", rootCmd.PersistentFlags().Lookup("account"))
	rootCmd.MarkPersistentFlagRequired("account")
	rootCmd.PersistentFlags().StringVar(&salesChannel, "salesChannel", "", "Sales Channel")
	viper.BindPFlag("salesChannel", rootCmd.PersistentFlags().Lookup("salesChannel"))
	rootCmd.PersistentFlags().StringVar(&serviceCode, "serviceCode", "", "Service Code")
	viper.BindPFlag("serviceCode", rootCmd.PersistentFlags().Lookup("serviceCode"))
	rootCmd.MarkPersistentFlagRequired("serviceCode")
	rootCmd.PersistentFlags().StringVar(&incoterm, "incoterm", "", "incoterm")
	viper.BindPFlag("incoterm", rootCmd.PersistentFlags().Lookup("incoterm"))
	rootCmd.PersistentFlags().StringVar(&originCountry, "originCountry", "", "Origin Country")
	viper.BindPFlag("originCountry", rootCmd.PersistentFlags().Lookup("originCountry"))
	rootCmd.PersistentFlags().StringVar(&weightMin, "weightMin", "", "Lower threshold for weight warning (per unit)")
	viper.BindPFlag("weightMin", rootCmd.PersistentFlags().Lookup("weightMin"))
	rootCmd.PersistentFlags().StringVar(&weightMax, "weightMax", "", "Upper threshold for weight warning (per unit)")
	viper.BindPFlag("weightMax", rootCmd.PersistentFlags().Lookup("weightMax"))

	rootCmd.PersistentFlags().StringVar(&description, "description", "", "description")
	viper.BindPFlag("description", rootCmd.PersistentFlags().Lookup("description"))
	rootCmd.MarkPersistentFlagRequired("description")
	rootCmd.PersistentFlags().StringVar(&HSCode, "HSCode", "", "HSCode")
	viper.BindPFlag("HSCode", rootCmd.PersistentFlags().Lookup("HSCode"))
	rootCmd.PersistentFlags().StringVar(&unitWeightStr, "unitWeight", "", "unitWeight")
	viper.BindPFlag("unitWeight", rootCmd.PersistentFlags().Lookup("unitWeight"))
	rootCmd.MarkPersistentFlagRequired("unitWeight")

	rootCmd.AddCommand(convertCmd)
	// Here you will define your flags and configuration settings.

	// Cobra supports Persistent Flags which will work for this command
	// and all subcommands, e.g.:
	// convertCmd.PersistentFlags().String("foo", "", "A help for foo")

	// Cobra supports local flags which will only run when this command
	// is called directly, e.g.:
	// convertCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
}
