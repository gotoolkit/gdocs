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

	"encoding/json"
	"github.com/spf13/cobra"
	"google.golang.org/api/sheets/v4"
	"io/ioutil"
	"log"
	"strings"
	"strconv"
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

		if jsonFile == "" {
			log.Fatalf("Need set json file use -j")
		}

		if len(readRanges) == 0 {
			log.Fatalf("Need set read ranges -s")
		}

		resp, err := service.Spreadsheets.Values.BatchGet(sheetId).Ranges(readRanges...).Do()
		if err != nil {
			log.Fatalf("Unable to retrieve data from sheet. %v", err)
		}

		if len(resp.ValueRanges) > 0 {
			rFace := make(map[string]interface{})

			for column, row := range resp.ValueRanges {
				if column == 0 {
					previousKey := ""
					//previous2Key := ""
					var r2Face map[string]interface{}
					//var r3Face map[string]interface{}
					for _, v := range row.Values {

						keys := strings.Split(v[0].(string), ".")

						converKeysToMap(previousKey, keys, rFace, r2Face)

					}
				}
				//fmt.Println(rFace)
			}
			jData, _ := json.Marshal(rFace)
			err = ioutil.WriteFile("output.json", jData, 0644)
		} else {
			fmt.Print("No data found.")
		}
	},
}

func converKeysToMap(previousKey string, keys []string, in map[string]interface{}, out map[string]interface{}) {
	if len(keys) > 1 {
		if previousKey != keys[0] {
			out = make(map[string]interface{})
			previousKey = keys[0]
			in[keys[0]] = out
		}

		out[keys[1]] = ""
	}

	index, err := strconv.Atoi(keys[0])
	if err != nil {
		out["index"] = index
	}

	//if previous2Key != keys[1] {
	//	r3Face = make(map[string]interface{})
	//	previousKey = keys[1]
	//	r2Face[keys[1]] = r3Face
	//}

	//keys = append(keys[:0], keys[:1]...)

	fmt.Println(keys)
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
