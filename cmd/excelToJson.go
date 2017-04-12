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

		if len(readRanges) == 0 {
			log.Fatalf("Need set read ranges -s")
		}

		resp, err := service.Spreadsheets.Values.BatchGet(sheetId).Ranges(readRanges...).Do()
		if err != nil {
			log.Fatalf("Unable to retrieve data from sheet. %v", err)
		}

		if len(resp.ValueRanges) > 0 {

			jKeys := resp.ValueRanges[0].Values
			jValues := resp.ValueRanges[1].Values
			jsonMap := make(map[string]string, len(jKeys))
			for i := 0; i < len(jKeys); i++ {
				jsonMap[jKeys[i][0].(string)] = jValues[i][0].(string)
			}
			rFace := make(map[string]interface{})
			for key, value := range jsonMap {
				keys := strings.Split(key, ".")
				convertToMap(rFace, keys, value)
			}

			jData, _ := json.Marshal(rFace)
			err = ioutil.WriteFile("output.json", jData, 0644)
		} else {
			fmt.Print("No data found.")
		}
	},
}

func convertToMap(in interface{}, keys []string, value string) {
	if len(keys) == 1 {
		in.(map[string]interface{})[keys[0]] = value
		return
	}

	val, ok := in.(map[string]interface{})[keys[0]]
	if !ok {
		val = make(map[string]interface{})
		in.(map[string]interface{})[keys[0]] = val
	}

	keys = append(keys[:0], keys[1:]...)
	convertToMap(val, keys, value)
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
