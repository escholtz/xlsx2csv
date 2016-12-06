package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"os"
	"path"
	"strings"

	"github.com/tealeg/xlsx"
)

func convert(sheet *xlsx.Sheet, name string) error {
	f, err := os.Create(name)
	if err != nil {
		return err
	}
	defer f.Close()

	w := csv.NewWriter(f)
	defer w.Flush()

	var cells []string
	for _, r := range sheet.Rows {
		if len(cells) != len(r.Cells) {
			cells = make([]string, len(r.Cells))
		}

		for i, c := range r.Cells {
			s, _ := c.String()
			cells[i] = s
		}
		w.Write(cells)
	}
	return nil
}

func main() {
	flag.Parse()
	args := flag.Args()
	if len(args) == 0 {
		fmt.Println("usage: xlsx2csv [filename ...]")
		os.Exit(1)
	}

	for _, fn := range args {
		f, err := xlsx.OpenFile(fn)
		if err != nil {
			fmt.Println(err)
			os.Exit(1)
		}

		if len(f.Sheets) == 0 {
			fmt.Println("%s no sheets", fn)
			os.Exit(1)
		}

		sheet := f.Sheets[0]

		if len(f.Sheets) > 1 {
			msg := ("warn: %s want 1 sheet found %d, " +
				"choosing longest sheet\n")
			fmt.Printf(msg, fn, len(f.Sheets))

			for i, s := range f.Sheets {
				msg := "  sheet %d, rows %d\n"
				fmt.Printf(msg, i, len(s.Rows))
				if len(s.Rows) > len(sheet.Rows) {
					sheet = s
				}
			}
		}

		fOut := strings.TrimSuffix(fn, path.Ext(fn)) + ".csv"
		err = convert(sheet, fOut)
		if err != nil {
			fmt.Println(err)
			os.Exit(1)
		}
	}
}
