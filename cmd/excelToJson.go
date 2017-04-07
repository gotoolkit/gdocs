// Copyright © 2017 NAME HERE <EMAIL ADDRESS>
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
	"google.golang.org/api/sheets/v4"
	"log"
)

// excelToJsonCmd represents the excelToJson command
var excelToJsonCmd = &cobra.Command{
	Use:   "excelToJson",
	Short: "A brief description of your command",
	Long: `A longer description that spans multiple lines and likely contains examples
and usage of using your command. For example:

Cobra is a CLI library for Go that empowers applications.
This application is a tool to generate the needed files
to quickly create a Cobra application.`,
	Run: func(cmd *cobra.Command, args []string) {
		client = initClient()
		service, err := sheets.New(client)
		if err != nil {
			log.Fatalf("Unable to retrieve Sheets Client %v", err)
		}

		if sheetId == "" {
			log.Fatalf("Need set sheet id use -i")
		}

		if readRange == "" {
			log.Fatalf("Need set read range use -r")
		}

		if jsonFile == "" {
			log.Fatalf("Need set json file use -j")
		}

		ranges := []string{"A1:A300", "C1:C300"} // TODO: Update placeholder value.

		resp, err := service.Spreadsheets.Values.BatchGet(sheetId).Ranges(ranges...).Do()
		if err != nil {
			log.Fatalf("Unable to retrieve data from sheet. %v", err)
		}

		if len(resp.ValueRanges) > 0 {
			fmt.Println("Name, Major:")
			for column, row := range resp.ValueRanges {
				// Print columns A and E, which correspond to indices 0 and 4.
				fmt.Println(column , row)
			}
		} else {
			fmt.Print("No data found.")
		}
	},
}

func init() {
	RootCmd.AddCommand(excelToJsonCmd)

	// Here you will define your flags and configuration settings.

	// Cobra supports Persistent Flags which will work for this command
	// and all subcommands, e.g.:
	// excelToJsonCmd.PersistentFlags().String("foo", "", "A help for foo")

	// Cobra supports local flags which will only run when this command
	// is called directly, e.g.:
	// excelToJsonCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
}
