package main

import (
	"encoding/json"
	"fmt"
	"net/http"
	"os"
	"strconv"
	"strings"
	"sync"
	"time"
)

type Comics struct {
	Month      string `json:"month"`
	Num        int    `json:"num"`
	Link       string `json:"link"`
	Year       string `json:"year"`
	News       string `json:"news"`
	SafeTitle  string `json:"safe_title"`
	Transcript string `json:"transcript"`
	Alt        string `json:"alt"`
	Img        string `json:"img"`
	Title      string `json:"title"`
	Day        string `json:"day"`
	comicLink  string
}

var wg sync.WaitGroup

func main() {

	var builder strings.Builder
	ch := make(chan map[int]string, 2000)

	for i := 1; i <= 100; i++ {
		go comicInfo("https://xkcd.com/"+strconv.Itoa(i)+"/info.0.json", i, ch)
		wg.Add(1)
	}

	wg.Wait()
	close(ch)

	for el := range ch {

	}

	file, _ := os.Create("d:\\comics.txt")
	file.WriteString(builder.String())
	file.Close()
}

func comicInfo(link string, i int, ch chan map[int]string) {
	try := 1
again:
	resp, _ := http.Get(link)

	if resp.StatusCode != http.StatusOK {
		resp.Body.Close()

		if resp.StatusCode == 503 {
			fmt.Println(i, "=", resp.Status, "| Try:", try)
			time.Sleep(time.Duration(try) * 500 * time.Millisecond)
			try++
			if try < 6 {
				goto again
			}
		}

		fmt.Println(i, "=", resp.Status)
		ch <- map[int]string{i: "Error in response"}
	}

	var com Comics

	if err := json.NewDecoder(resp.Body).Decode(&com); err != nil {
		fmt.Println(i, " = ", err)
		ch <- map[int]string{i: "Error in decode"}
	}

	com.comicLink = "https://xkcd.com/" + strconv.Itoa(i) + "/"
	fmt.Println("Finished with", i)

	ch <- map[int]string{i: "\t\t\t\t\t\t\tComic #" + strconv.Itoa(i) + ": " + com.comicLink + "\n\n" + com.Transcript + "\n\n"}

	defer wg.Done()
}
