package main

import (
	"encoding/json"
	"errors"
	"fmt"
	"github.com/gorpher/winspool-cgo/model"
	"log"
	"os"
	"os/signal"
	"strconv"
	"syscall"

	"github.com/cheynewallace/tabby"
	"github.com/gorpher/gone"
	"github.com/gorpher/winspool-cgo/lib"
	"github.com/gorpher/winspool-cgo/winspool"
	cli "github.com/urfave/cli/v2"
)

var (
	version  = "nil"
	hash     = "nil"
	datetime = "nil"
)

type App struct {
	spool *winspool.WinSpool
	jobs  chan *lib.Job
}

func (a *App) ListPrinter(c *cli.Context) error {
	printers, err := a.spool.GetPrinters()
	if err != nil {
		return err
	}
	OutputPrintList(printers)
	return nil
}

func (a *App) InspectPrinter(c *cli.Context) error {
	args := c.Args()
	if args.Len() < 1 {
		return errors.New("请输入打印机名称")
	}
	printerName := args.Get(0)
	printers, err := a.spool.GetPrinters()
	if err != nil {
		return errors.New("没有可用打印机")
	}
	var printer *lib.Printer
	for _, p := range printers {
		if p.Name == printerName {
			printer = &p
		}
	}
	if printer == nil {
		return errors.New("打印机不存在")
	}
	body, err := json.MarshalIndent(*printer, "", "   ")
	if err != nil {
		return err
	}
	fmt.Println(string(body))
	return nil
}

func (a *App) AddJob(c *cli.Context) error {
	filename := c.String("filename")
	if filename == "" {
		return errors.New("文件名不能为空")
	}
	printerName := c.String("printer")
	if printerName == "" {
		return errors.New("打印机不能为空")
	}
	if !gone.FileExist(filename) {
		return fmt.Errorf("文件 %s 不存在", filename)
	}
	printers, err := a.spool.GetPrinters()
	if err != nil {
		return errors.New("没有可用打印机")
	}
	var printer *lib.Printer
	for _, p := range printers {
		if p.Name == printerName {
			printer = &p
		}
	}
	if printer == nil {
		return errors.New("打印机不存在")
	}
	var nativeJobQueueSize uint = 2
	printer.NativeJobSemaphore = lib.NewSemaphore(nativeJobQueueSize)

	jobID, err := a.spool.Print(printer, filename, gone.RandLower(8), &model.JobTicket{
		Copies: &model.CopiesTicketItem{
			Copies: 1,
		},
	})
	if err != nil {
		return err
	}
	fmt.Printf("{\"job_id\":%d}\n", jobID)
	return nil
}

func (a *App) StatusJob(c *cli.Context) error {
	fmt.Println("查看打印机job状态")
	args := c.Args()
	if args.Len() < 2 {
		return errors.New("usage state <printerName> <jobID>")
	}
	printerName := args.Get(0)
	printers, err := a.spool.GetPrinters()
	if err != nil {
		return errors.New("没有可用打印机")
	}
	var printer *lib.Printer
	for _, p := range printers {
		if p.Name == printerName {
			printer = &p
		}
	}
	if printer == nil {
		return errors.New("打印机不存在")
	}
	jobIDStr := args.Get(1)
	jobID, err := strconv.ParseUint(jobIDStr, 10, 32)
	if err != nil {
		return errors.New("jobID 错误")
	}

	state, err := a.spool.GetJobState(printerName, uint32(jobID))
	if err != nil {
		return err
	}

	body, err := json.MarshalIndent(state, "", "   ")
	if err != nil {
		return err
	}
	fmt.Println(string(body))
	return nil
}

func (a *App) ListJob(c *cli.Context) error {
	fmt.Println("查看打印机作业列表")
	args := c.Args()
	if args.Len() < 1 {
		return errors.New("usage state <printerName>")
	}
	printerName := args.Get(0)
	list, err := a.spool.JobList(printerName)
	if err != nil {
		return err
	}
	OutputJobList(list)
	return nil
}

func (a *App) Version(c *cli.Context) error {
	fmt.Printf("echo-service has version %s built from %s on %s\n", version, hash, datetime)
	return nil
}

func NewApp() *cli.App {
	spool, err := winspool.NewWinSpool()
	if err != nil {
		panic(err)
	}
	jobs := make(chan *lib.Job, 10)
	app := &App{
		spool: spool,
		jobs:  jobs,
	}

	return &cli.App{
		Name:  "printpdf",
		Usage: "打印机操作命令行程序",
		Commands: []*cli.Command{
			{
				Name:   "version",
				Action: app.Version,
				Usage:  "查看版本号",
			},
			{
				Name:  "printer",
				Usage: "打印机操作",
				Subcommands: []*cli.Command{
					{
						Name:   "ls",
						Usage:  "获取打印机列表",
						Action: app.ListPrinter,
					},
					{
						Name:   "inspect",
						Usage:  "获取打印机详情",
						Action: app.InspectPrinter,
					},
				},
			},
			{
				Name:  "job",
				Usage: "作业",
				Subcommands: []*cli.Command{
					{
						Flags: []cli.Flag{
							&cli.StringFlag{
								Name:    "filename",
								Aliases: []string{"f"},
								Usage:   "文件路径",
							},
							&cli.StringFlag{
								Name:    "printer",
								Aliases: []string{"p"},
								Usage:   "打印机名称",
							},
						},
						Name:   "add",
						Usage:  "添加打印作业",
						Action: app.AddJob,
					},
					{

						Name:   "status",
						Usage:  "打印作业状态",
						Action: app.StatusJob,
					},
					{

						Name:   "ls",
						Usage:  "打印机作业列表",
						Action: app.ListJob,
					},
				},
			},
			// ===========================
			{
				Name:   "printers",
				Usage:  "获取打印机列表",
				Action: app.ListPrinter,
			},
		},
		After: func(context *cli.Context) error {
			//printerManager.Quit()
			return nil
		},
	}
}

// Blocks until Ctrl-C or SIGTERM.
func waitIndefinitely() {
	ch := make(chan os.Signal)
	signal.Notify(ch, os.Interrupt, syscall.SIGTERM)
	<-ch

	go func() {
		// In case the process doesn't die quickly, wait for a second termination request.
		<-ch
		fmt.Println("Second termination request received")
		os.Exit(1)
	}()
}

func main() {
	err := NewApp().Run(os.Args)
	if err != nil {
		log.Fatal(err)
	}
}

func OutputPrintList(printers []lib.Printer) {
	t := tabby.New()
	t.AddHeader("名称", "名称2", "驱动", "状态")
	for _, printer := range printers {
		t.AddLine(printer.Name, printer.DefaultDisplayName, printer.Model, printer.State.State)
	}
	t.Print()
}

func OutputJobList(jobs []winspool.Job) {
	t := tabby.New()
	t.AddHeader("作业ID", "打印机名称", "打印类型", "状态")
	for _, printer := range jobs {
		t.AddLine(printer.JobID, printer.PrinterName, printer.Datatype, printer.Status)
	}
	t.Print()
}
