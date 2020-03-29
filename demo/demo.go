package main

import (
	"flag"

	xlst "github.com/ivahaev/go-xlsx-templater"
)

func main() {
	var (
		in  string
		out string
	)

	flag.StringVar(&in, "in", "template.xlsx", "path to the template")
	flag.StringVar(&out, "out", "report.xlsx", "path to result file")

	flag.Parse()

	doc := xlst.New()
	doc.ReadTemplate(in)

	ctx := map[string]interface{}{
		"name":           "Github User",
		"groupHeader":    "Group name",
		"nameHeader":     "Item name",
		"quantityHeader": "Quantity",
		"groups": []map[string]interface{}{
			{
				"name":  "Work",
				"total": 3,
				"items": []map[string]interface{}{
					{
						"name":     "Pen",
						"quantity": 2,
					},
					{
						"name":     "Pencil",
						"quantity": 1,
					},
				},
			},
			{
				"name":      "Weekend",
				"total":     36,
				"important": true,
				"items": []map[string]interface{}{
					{
						"name":     "Condom",
						"quantity": 12,
					},
					{
						"name":     "Beer",
						"quantity": 24,
					},
				},
			},
		},
	}

	err := doc.Render(ctx)
	if err != nil {
		panic(err)
	}

	err = doc.Save(out)
	if err != nil {
		panic(err)
	}
}
