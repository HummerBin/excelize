package main

/*
完成对读取excel数据、执行接口请求、写入新的excel
*/
import (
	"bytes"
	"encoding/json"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io/ioutil"
	"net/http"
	"runtime"
	"runtime/debug"
	"strconv"
	"sync"
	"time"
)

const xd_key = "LKdeBLwAdeJSHIURghtwehrjtOV"
const xj_key = "LKdeBLwAdeJSHIURghtwehrjtOV"

func main() {
	runtime.GOMAXPROCS(runtime.NumCPU())
	t1 := time.Now()
	fmt.Println("执行开始时间", t1)
	fileArray := [...]string{"./bank.xlsx"}
	var wg sync.WaitGroup
	wg.Add(len(fileArray))
	for i := 0; i < len(fileArray); i++ {
		go func(fileName string) {
			var ma = make([]map[string]string, 0)
			fmt.Println("文件开始处理", fileName)
			xlsx, err := excelize.OpenFile(fileName)
			if err != nil {
				fmt.Println(err)
				return
			}
			rows := xlsx.GetRows("Sheet1")
			rownums := len(rows)
			fmt.Println(rownums)
			for j := 1; j <= rownums; j++ {
				fmt.Println("正在处理第" + strconv.Itoa(j) + "条数据")
				cell := xlsx.GetCellValue("Sheet1", "A"+strconv.Itoa(j))
				cellName := xlsx.GetCellValue("Sheet1", "B"+strconv.Itoa(j))
				if cell != "" {
					httpPostJson(cell, cellName, &ma)
				}
			}
			//export excel
			exportExcel(ma, fileName+"temp.xlsx")
			fmt.Println("文件处理完成", fileName)
			wg.Done()
		}(fileArray[i])
	}
	wg.Wait()
	fmt.Println("执行花费", time.Now().Sub(t1))
}

//写入excel
func exportExcel(resultData []map[string]string, fileName string) {
	xlsx1 := excelize.NewFile()
	index := xlsx1.NewSheet("Sheet1")
	for i, v := range resultData {
		fmt.Println(v)
		xlsx1.SetCellValue("Sheet1", "A"+strconv.Itoa(i+1), v["userid"])
		xlsx1.SetCellValue("Sheet1", "B"+strconv.Itoa(i+1), v["mobile"])
		xlsx1.SetCellValue("Sheet1", "C"+strconv.Itoa(i+1), v["message"])
	}
	xlsx1.SetActiveSheet(index)
	// Save xlsx file by the given path.
	err1 := xlsx1.SaveAs(fileName)
	if err1 != nil {
		fmt.Println(err1)
	}
}

//发起posti 请求
func httpPostJson(userid string, mobile string, ma *[]map[string]string) *[]map[string]string {

	url := "http://commonservices.zhphfinance.com/zhph_commonServices/webservice/personal/updateCgtMobile"
	//u := PostDate{tdUserName: "td_zhph_xfjr", userId: userid,deviceType:"PC"}
	song := make(map[string]interface{})
	data := make(map[string]interface{})
	data["userId"] = userid
	data["mobile"] = mobile
	song["tdUserName"] = "td_zhph_xd"
	song["key"] = xd_key
	song["data"] = data
	bytesData, err := json.Marshal(song)
	req1 := bytes.NewBuffer([]byte(bytesData))
	req, err := http.NewRequest("POST", url, req1)
	req.Header.Set("Content-Type", "application/json")

	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		fmt.Println(err)
		fmt.Println(debug.Stack())
	}
	defer resp.Body.Close()

	statuscode := resp.StatusCode
	hea := resp.Header
	body, _ := ioutil.ReadAll(resp.Body)
	var dat map[string]interface{}
	if err := json.Unmarshal([]byte(body), &dat); err == nil {
		fmt.Println(dat)
		fmt.Println(dat["data"])
		mapV := make(map[string]string)
		mapV["userid"] = userid
		mapV["mobile"] = mobile
		mapV["message"] = dat["message"].(string)
		*ma = append(*ma, mapV)
	} else {
		fmt.Println(err)
	}

	fmt.Println(string(body))
	fmt.Println(statuscode)
	fmt.Println(hea)
	return ma
}
