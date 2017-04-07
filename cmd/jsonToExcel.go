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
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"

	"sort"
	"github.com/spf13/cobra"
	"google.golang.org/api/sheets/v4"
)

// jsonToExcelCmd represents the jsonToExcel command
var jsonToExcelCmd = &cobra.Command{
	Use:   "jsonToExcel",
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

		jsonData, err := ioutil.ReadFile(jsonFile)
		if err != nil {
			log.Fatalf("Unable to read json file %v", err)
		}

		jFace := make(map[string]interface{})

		err = json.Unmarshal(jsonData, &jFace)

		if err != nil {
			log.Fatalf("Unable to parse json file %v", err)
		}
		rFace := make(map[string]interface{})

		jsonToKeyValue("", jFace, rFace)

		values := make([][]interface{}, len(rFace))
		var keys []string
		for k := range rFace {
			keys = append(keys, k)
		}
		sort.Strings(keys)

		for k, v := range keys {
			cells := make([]interface{}, 2)
			cells[0] = v
			cells[1] = values[k]
			values = append(values, cells)
		}
		valueRange := &sheets.ValueRange{
			Values:         values,
			MajorDimension: "ROWS",
		}
		_, err = service.Spreadsheets.Values.Update(sheetId, readRange, valueRange).ValueInputOption("RAW").Do()
		if err != nil {
			log.Fatalf("Unable to update data from sheet. %v", err)
		}
	},
}

func jsonToKeyValue(pref string, in map[string]interface{}, out map[string]interface{}) {
	if pref != "" {
		pref += "."
	}
	for key, value := range in {
		switch v := value.(type) {
		case int:
			out[pref+key] = v
		case float64:
			out[pref+key] = v
		case string:
			out[pref+key] = v
		case bool:
			out[pref+key] = v
		case []interface{}:
			iArray := value.([]interface{})
			for i := 0; i < len(iArray); i++ {
				iMap := make(map[string]interface{})
				iMap[fmt.Sprint(i)] = iArray[i]
				jsonToKeyValue(fmt.Sprint(pref, key), iMap, out)
			}
		default:
			jsonToKeyValue(pref+key, value.(map[string]interface{}), out)
		}
	}
}

func init() {
	RootCmd.AddCommand(jsonToExcelCmd)

	// Here you will define your flags and configuration settings.

	// Cobra supports Persistent Flags which will work for this command
	// and all subcommands, e.g.:
	// jsonToExcelCmd.PersistentFlags().String("foo", "", "A help for foo")

	// Cobra supports local flags which will only run when this command
	// is called directly, e.g.:
	// jsonToExcelCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")

}
