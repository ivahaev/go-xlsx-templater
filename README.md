# go-xlsx-templater
Simple .xlsx (Excel XML document) template based document generator

## Installation

```
    go get -u "github.com/ivahaev/go-xlsx-templater"
```


## Usage

### Importing to your project

```go
    import "github.com/ivahaev/go-xlsx-templater"
```

#### type Xlst

```go
type Xlst struct {
    // contains filtered or unexported fields
}
```

Xlst Represents template struct

#### func  New

```go
func New() *Xlst
```
New() creates new Xlst struct and returns pointer to it

#### func (*Xlst) ReadTemplate

```go
func (m *Xlst) ReadTemplate(path string) error
```
ReadTemplate() reads template from disk and stores it in a struct

#### func (*Xlst) Render

```go
func (m *Xlst) Render(ctx map[string]interface{}) error
```
Render() renders report and stores it in a struct

#### func (*Xlst) Save

```go
func (m *Xlst) Save(path string) error
```
Save() saves generated report to disk

#### func (*Xlst) Write

```go
func (m *Xlst) Write(writer io.Writer) error
```
Write() writes generated report to provided writer
