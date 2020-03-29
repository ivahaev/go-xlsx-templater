# go-xlsx-templater

Simple **.xlsx** (Excel XML document) template based document generator using handlebars.

Takes input **.xlsx** documents with mustache sippets in it and renders new document with snippets replaced by provided context.

Thanks to `github.com/tealeg/xlsx` and `github.com/aymerick/raymond` for useful libs.

[About package in russian](http://ivahaev.ru/go-xlsx-templater/)

[Godoc](https://godoc.org/github.com/ivahaev/go-xlsx-templater)

## Installation

```console
    go get -u "github.com/ivahaev/go-xlsx-templater"
```

## Usage

### Import to your project

```go
    import "github.com/ivahaev/go-xlsx-templater"
```

### Prepare **template.xlsx** template

Filename may be any of course. For slices use dot notation `{{items.name}}`. When parser meets dot notation it will repeats contains row.

#### Blocks

For complex manipulations with rows you can use block notaion. Each block begins with `{{blockName fieldName}}` and ends with `{{end}}`. They must be defined in the first cell of the row. All other content of the row will be ignored. Currently supported block types:

- **range** - allows you to repeat several rows for each element in the `fieldName`. Of course, `fieldName` must contain a slice or an array.
- **if** - allows you to display rows conditionally. `fieldName` will be considered as a *true* condition, if it exists in the context and doesn't contain default value of its type. You also can use `{{else}}` for displaying rows when condition is *false*.

![Sample document image](./template.png)

### Prepare context data

```go
    ctx := map[string]interface{}{
        "name": "Github User",
        "groupHeader": "Group name",
        "nameHeader": "Item name",
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
                "name":  "Weekend",
                "total": 36,
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
```

### Read template, render with context and save to disk

Error processing omited in example.

```go
    doc := xlst.New()
    doc.ReadTemplate("./template.xlsx")
    doc.Render(ctx)
    doc.Save("./report.xlsx")
```

### Enjoy created report

![Report image](./report.png)

## Documentation

### type Xlst

```go
type Xlst struct {
    // contains filtered or unexported fields
}
```

Xlst Represents template struct

### func  New

```go
func New() *Xlst
```

New() creates new Xlst struct and returns pointer to it

### func (*Xlst) ReadTemplate

```go
func (m *Xlst) ReadTemplate(path string) error
```

ReadTemplate() reads template from disk and stores it in a struct

### func (*Xlst) Render

```go
func (m *Xlst) Render(ctx map[string]interface{}) error
```

Render() renders report and stores it in a struct

### func (*Xlst) Save

```go
func (m *Xlst) Save(path string) error
```

Save() saves generated report to disk

### func (*Xlst) Write

```go
func (m *Xlst) Write(writer io.Writer) error
```

Write() writes generated report to provided writer
