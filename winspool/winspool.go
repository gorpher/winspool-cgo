// Copyright 2015 Google Inc. All rights reserved.

// Use of this source code is governed by a BSD-style
// license that can be found in the LICENSE file or at
// https://developers.google.com/open-source/licenses/bsd

//go:build windows
// +build windows

package winspool

import (
	"errors"
	"github.com/gorpher/winspool-cgo/lib"
	"github.com/gorpher/winspool-cgo/model"
	"golang.org/x/sys/windows"
	"strconv"
	"strings"
)

// winspoolPDS represents capabilities that WinSpool always provides.
var winspoolPDS = model.PrinterDescriptionSection{
	SupportedContentType: &[]model.SupportedContentType{
		model.SupportedContentType{ContentType: "application/pdf"},
	},
	FitToPage: &model.FitToPage{
		Option: []model.FitToPageOption{
			model.FitToPageOption{
				Type:      model.FitToPageNoFitting,
				IsDefault: true,
			},
			model.FitToPageOption{
				Type:      model.FitToPageFitToPage,
				IsDefault: false,
			},
		},
	},
}

// WinSpool Interface between Go and the Windows API.
type WinSpool struct {
}

func NewWinSpool() (*WinSpool, error) {
	ws := WinSpool{}
	return &ws, nil
}

func convertPrinterState(wsStatus uint32, wsAttributes uint32) *model.PrinterStateSection {
	state := model.PrinterStateSection{
		State:       model.CloudDeviceStateIdle,
		VendorState: &model.VendorState{},
	}

	if wsStatus&(PRINTER_STATUS_PRINTING|PRINTER_STATUS_PROCESSING) != 0 {
		state.State = model.CloudDeviceStateProcessing
	}

	if wsStatus&PRINTER_STATUS_PAUSED != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateWarning,
			DescriptionLocalized: model.NewLocalizedString("printer paused"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_ERROR != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("printer error"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_PENDING_DELETION != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("printer is being deleted"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_PAPER_JAM != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("paper jam"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_PAPER_OUT != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("paper out"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_MANUAL_FEED != 0 {
		vs := model.VendorStateItem{
			State:                model.VendorStateInfo,
			DescriptionLocalized: model.NewLocalizedString("manual feed mode"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_PAPER_PROBLEM != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("paper problem"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}

	// If PRINTER_ATTRIBUTE_WORK_OFFLINE is set
	// spooler won't despool any jobs to the printer.
	// At least for some USB printers, this flag is controlled
	// automatically by the system depending on the state of physical connection.
	if wsStatus&PRINTER_STATUS_OFFLINE != 0 || wsAttributes&PRINTER_ATTRIBUTE_WORK_OFFLINE != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("printer is offline"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_IO_ACTIVE != 0 {
		vs := model.VendorStateItem{
			State:                model.VendorStateInfo,
			DescriptionLocalized: model.NewLocalizedString("active I/O state"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_BUSY != 0 {
		vs := model.VendorStateItem{
			State:                model.VendorStateInfo,
			DescriptionLocalized: model.NewLocalizedString("busy"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_OUTPUT_BIN_FULL != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("output bin is full"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_NOT_AVAILABLE != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("printer not available"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_WAITING != 0 {
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("waiting"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_INITIALIZING != 0 {
		vs := model.VendorStateItem{
			State:                model.VendorStateInfo,
			DescriptionLocalized: model.NewLocalizedString("intitializing"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_WARMING_UP != 0 {
		vs := model.VendorStateItem{
			State:                model.VendorStateInfo,
			DescriptionLocalized: model.NewLocalizedString("warming up"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_TONER_LOW != 0 {
		vs := model.VendorStateItem{
			State:                model.VendorStateWarning,
			DescriptionLocalized: model.NewLocalizedString("toner low"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_NO_TONER != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("no toner"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_PAGE_PUNT != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("cannot print the current page"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_USER_INTERVENTION != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("user intervention required"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_OUT_OF_MEMORY != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("out of memory"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_DOOR_OPEN != 0 {
		state.State = model.CloudDeviceStateStopped
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("door open"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_SERVER_UNKNOWN != 0 {
		vs := model.VendorStateItem{
			State:                model.VendorStateError,
			DescriptionLocalized: model.NewLocalizedString("printer status unknown"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}
	if wsStatus&PRINTER_STATUS_POWER_SAVE != 0 {
		vs := model.VendorStateItem{
			State:                model.VendorStateInfo,
			DescriptionLocalized: model.NewLocalizedString("power save mode"),
		}
		state.VendorState.Item = append(state.VendorState.Item, vs)
	}

	if len(state.VendorState.Item) == 0 {
		state.VendorState = nil
	}

	return &state
}

func getManModel(driverName string) (man string, model string) {
	man = "Google"
	model = "Cloud Printer"

	parts := strings.SplitN(driverName, " ", 2)
	if len(parts) > 0 && len(parts[0]) > 0 {
		man = parts[0]
	}
	if len(parts) > 1 && len(parts[1]) > 0 {
		model = parts[1]
	}

	return
}

// GetPrinters gets all Windows printers found on this computer.
func (ws *WinSpool) GetPrinters() ([]lib.Printer, error) {
	pi2s, err := EnumPrinters2()
	if err != nil {
		return nil, err
	}

	printers := make([]lib.Printer, 0, len(pi2s))
	for _, pi2 := range pi2s {
		printerName := pi2.GetPrinterName()
		portName := pi2.GetPortName()
		devMode := pi2.GetDevMode()

		manufacturer, model1 := getManModel(pi2.GetDriverName())
		printer := lib.Printer{
			Name:               printerName,
			DefaultDisplayName: printerName,
			Manufacturer:       manufacturer,
			Model:              model1,
			State:              convertPrinterState(pi2.GetStatus(), pi2.GetAttributes()),
			Description:        &model.PrinterDescriptionSection{},
			Tags: map[string]string{
				"printer-location": pi2.GetLocation(),
			},
		}

		// Advertise color based on default value, which should be a solid indicator
		// of color-ness, because the source of this devMode object is EnumPrinters.
		if def, ok := devMode.GetColor(); ok {
			if def == DMCOLOR_COLOR {
				printer.Description.Color = &model.Color{
					Option: []model.ColorOption{
						model.ColorOption{
							VendorID:                   strconv.FormatInt(int64(DMCOLOR_COLOR), 10),
							Type:                       model.ColorTypeStandardColor,
							IsDefault:                  true,
							CustomDisplayNameLocalized: model.NewLocalizedString("Color"),
						},
						model.ColorOption{
							VendorID:                   strconv.FormatInt(int64(DMCOLOR_MONOCHROME), 10),
							Type:                       model.ColorTypeStandardMonochrome,
							IsDefault:                  false,
							CustomDisplayNameLocalized: model.NewLocalizedString("Monochrome"),
						},
					},
				}
			} else if def == DMCOLOR_MONOCHROME {
				printer.Description.Color = &model.Color{
					Option: []model.ColorOption{
						model.ColorOption{
							VendorID:                   strconv.FormatInt(int64(DMCOLOR_MONOCHROME), 10),
							Type:                       model.ColorTypeStandardMonochrome,
							IsDefault:                  true,
							CustomDisplayNameLocalized: model.NewLocalizedString("Monochrome"),
						},
					},
				}
			}
		}

		if def, ok := devMode.GetDuplex(); ok {
			duplex, err := DeviceCapabilitiesInt32(printerName, portName, DC_DUPLEX)
			if err != nil {
				return nil, err
			}
			if duplex == 1 {
				printer.Description.Duplex = &model.Duplex{
					Option: []model.DuplexOption{
						model.DuplexOption{
							Type:      model.DuplexNoDuplex,
							IsDefault: def == DMDUP_SIMPLEX,
						},
						model.DuplexOption{
							Type:      model.DuplexLongEdge,
							IsDefault: def == DMDUP_VERTICAL,
						},
						model.DuplexOption{
							Type:      model.DuplexShortEdge,
							IsDefault: def == DMDUP_HORIZONTAL,
						},
					},
				}
			}
		}

		if def, ok := devMode.GetOrientation(); ok {
			orientation, err := DeviceCapabilitiesInt32(printerName, portName, DC_ORIENTATION)
			if err != nil {
				return nil, err
			}
			if orientation == 90 || orientation == 270 {
				printer.Description.PageOrientation = &model.PageOrientation{
					Option: []model.PageOrientationOption{
						model.PageOrientationOption{
							Type:      model.PageOrientationPortrait,
							IsDefault: def == DMORIENT_PORTRAIT,
						},
						model.PageOrientationOption{
							Type:      model.PageOrientationLandscape,
							IsDefault: def == DMORIENT_LANDSCAPE,
						},
					},
				}
			}
		}

		if def, ok := devMode.GetCopies(); ok {
			copies, err := DeviceCapabilitiesInt32(printerName, portName, DC_COPIES)
			if err != nil {
				return nil, err
			}
			if copies > 1 {
				printer.Description.Copies = &model.Copies{
					Default: int32(def),
					Max:     copies,
				}
			}
		}

		printer.Description.MediaSize, err = convertMediaSize(printerName, portName, devMode)
		if err != nil {
			return nil, err
		}

		if def, ok := devMode.GetCollate(); ok {
			collate, err := DeviceCapabilitiesInt32(printerName, portName, DC_COLLATE)
			if err != nil {
				return nil, err
			}
			if collate == 1 {
				printer.Description.Collate = &model.Collate{
					Default: def == DMCOLLATE_TRUE,
				}
			}
		}

		printers = append(printers, printer)
	}

	return printers, nil
}

func convertMediaSize(printerName, portName string, devMode *DevMode) (*model.MediaSize, error) {
	defSize, defSizeOK := devMode.GetPaperSize()
	defLength, defLengthOK := devMode.GetPaperLength()
	defWidth, defWidthOK := devMode.GetPaperWidth()

	names, err := DeviceCapabilitiesStrings(printerName, portName, DC_PAPERNAMES, 64*2)
	if err != nil {
		return nil, err
	}
	papers, err := DeviceCapabilitiesUint16Array(printerName, portName, DC_PAPERS)
	if err != nil {
		return nil, err
	}
	sizes, err := DeviceCapabilitiesInt32Pairs(printerName, portName, DC_PAPERSIZE)
	if err != nil {
		return nil, err
	}
	if len(names) != len(papers) || len(names) != len(sizes)/2 {
		return nil, nil
	}

	ms := model.MediaSize{
		Option: make([]model.MediaSizeOption, 0, len(names)),
	}

	var foundDef bool
	for i := range names {
		if names[i] == "" {
			continue
		}
		// Convert from tenths-of-mm to micrometers
		width, length := sizes[2*i]*100, sizes[2*i+1]*100

		var def bool
		if !foundDef {
			if defSizeOK {
				if uint16(defSize) == papers[i] {
					def = true
					foundDef = true
				}
			} else if defLengthOK && int32(defLength) == length && defWidthOK && int32(defWidth) == width {
				def = true
				foundDef = true
			}
		}

		o := model.MediaSizeOption{
			Name:                       model.MediaSizeCustom,
			WidthMicrons:               width,
			HeightMicrons:              length,
			IsDefault:                  def,
			VendorID:                   strconv.FormatUint(uint64(papers[i]), 10),
			CustomDisplayNameLocalized: model.NewLocalizedString(names[i]),
		}
		ms.Option = append(ms.Option, o)
	}

	if !foundDef && len(ms.Option) > 0 {
		ms.Option[0].IsDefault = true
	}

	return &ms, nil
}

func convertJobState(wsStatus uint32) *model.JobState {
	var state model.JobState

	if wsStatus&(JOB_STATUS_SPOOLING|JOB_STATUS_PRINTING) != 0 {
		state.Type = model.JobStateInProgress

	} else if wsStatus&(JOB_STATUS_PRINTED|JOB_STATUS_COMPLETE) != 0 {
		state.Type = model.JobStateDone

	} else if wsStatus&JOB_STATUS_PAUSED != 0 || wsStatus == 0 {
		state.Type = model.JobStateDone

	} else if wsStatus&JOB_STATUS_ERROR != 0 {
		state.Type = model.JobStateAborted
		state.DeviceActionCause = &model.DeviceActionCause{model.DeviceActionCausePrintFailure}

	} else if wsStatus&(JOB_STATUS_DELETING|JOB_STATUS_DELETED) != 0 {
		state.Type = model.JobStateAborted
		state.UserActionCause = &model.UserActionCause{model.UserActionCauseCanceled}

	} else if wsStatus&(JOB_STATUS_OFFLINE|JOB_STATUS_PAPEROUT|JOB_STATUS_BLOCKED_DEVQ|JOB_STATUS_USER_INTERVENTION) != 0 {
		state.Type = model.JobStateStopped
		state.DeviceStateCause = &model.DeviceStateCause{model.DeviceStateCauseOther}

	} else {
		// Don't know what is going on. Get the job out of our queue.
		state.Type = model.JobStateAborted
		state.DeviceActionCause = &model.DeviceActionCause{model.DeviceActionCauseOther}
	}

	return &state
}

// GetJobState gets the current state of the job indicated by jobID.
func (ws *WinSpool) GetJobState(printerName string, jobID uint32) (*model.PrintJobStateDiff, error) {
	hPrinter, err := OpenPrinter(printerName)
	if err != nil {
		return nil, err
	}

	ji1, err := hPrinter.GetJob(int32(jobID))
	if err != nil {
		if err == ERROR_INVALID_PARAMETER {
			jobState := model.PrintJobStateDiff{
				State: &model.JobState{
					Type:              model.JobStateAborted,
					DeviceActionCause: &model.DeviceActionCause{model.DeviceActionCauseOther},
				},
			}
			return &jobState, nil
		}
		return nil, err
	}

	jobState := model.PrintJobStateDiff{
		State: convertJobState(ji1.GetStatus()),
	}
	return &jobState, nil
}

type jobContext struct {
	jobID    int32
	pDoc     PopplerDocument
	hPrinter HANDLE
	devMode  *DevMode
	hDC      HDC
	cSurface CairoSurface
	cContext CairoContext
}

func newJobContext(printerName, fileName, title string) (*jobContext, error) {
	pDoc, err := PopplerDocumentNewFromFile(fileName)
	if err != nil {
		return nil, err
	}
	hPrinter, err := OpenPrinter(printerName)
	if err != nil {
		pDoc.Unref()
		return nil, err
	}
	devMode, err := hPrinter.DocumentPropertiesGet(printerName)
	if err != nil {
		hPrinter.ClosePrinter()
		pDoc.Unref()
		return nil, err
	}
	err = hPrinter.DocumentPropertiesSet(printerName, devMode)
	if err != nil {
		hPrinter.ClosePrinter()
		pDoc.Unref()
		return nil, err
	}
	hDC, err := CreateDC(printerName, devMode)
	if err != nil {
		hPrinter.ClosePrinter()
		pDoc.Unref()
		return nil, err
	}
	jobID, err := hDC.StartDoc(title)
	if err != nil {
		hDC.DeleteDC()
		hPrinter.ClosePrinter()
		pDoc.Unref()
		return nil, err
	}
	hPrinter.SetJobUserName(jobID)
	cSurface, err := CairoWin32PrintingSurfaceCreate(hDC)
	if err != nil {
		hDC.EndDoc()
		hDC.DeleteDC()
		hPrinter.ClosePrinter()
		pDoc.Unref()
		return nil, err
	}
	cContext, err := CairoCreateContext(cSurface)
	if err != nil {
		cSurface.Destroy()
		hDC.EndDoc()
		hDC.DeleteDC()
		hPrinter.ClosePrinter()
		pDoc.Unref()
		return nil, err
	}
	c := jobContext{jobID, pDoc, hPrinter, devMode, hDC, cSurface, cContext}
	return &c, nil
}

func (c *jobContext) free() error {
	var err error
	err = c.cContext.Destroy()
	if err != nil {
		return err
	}
	err = c.cSurface.Destroy()
	if err != nil {
		return err
	}
	err = c.hDC.EndDoc()
	if err != nil {
		return err
	}
	err = c.hDC.DeleteDC()
	if err != nil {
		return err
	}
	err = c.hPrinter.ClosePrinter()
	if err != nil {
		return err
	}
	c.pDoc.Unref()
	return nil
}

func getScaleAndOffset(wDocPoints, hDocPoints float64, wPaperPixels, hPaperPixels, xMarginPixels, yMarginPixels, wPrintablePixels, hPrintablePixels, xDPI, yDPI int32, fitToPage bool) (scale, xOffsetPoints, yOffsetPoints float64) {

	wPaperPoints, hPaperPoints := float64(wPaperPixels*72)/float64(xDPI), float64(hPaperPixels*72)/float64(yDPI)

	var wPrintablePoints, hPrintablePoints float64
	if fitToPage {
		wPrintablePoints, hPrintablePoints = float64(wPrintablePixels*72)/float64(xDPI), float64(hPrintablePixels*72)/float64(yDPI)
	} else {
		wPrintablePoints, hPrintablePoints = wPaperPoints, hPaperPoints
	}

	xScale, yScale := wPrintablePoints/wDocPoints, hPrintablePoints/hDocPoints
	if xScale < yScale {
		scale = xScale
	} else {
		scale = yScale
	}

	xOffsetPoints = (wPaperPoints - wDocPoints*scale) / 2
	yOffsetPoints = (hPaperPoints - hDocPoints*scale) / 2

	return
}

func printPage(printerName string, i int, c *jobContext, fitToPage bool) error {
	pPage := c.pDoc.GetPage(i)
	defer pPage.Unref()

	if err := c.hPrinter.DocumentPropertiesSet(printerName, c.devMode); err != nil {
		return err
	}

	if err := c.hDC.ResetDC(c.devMode); err != nil {
		return err
	}

	// Set device to zero offset, and to points scale.
	xDPI := c.hDC.GetDeviceCaps(LOGPIXELSX)
	yDPI := c.hDC.GetDeviceCaps(LOGPIXELSY)
	xMarginPixels := c.hDC.GetDeviceCaps(PHYSICALOFFSETX)
	yMarginPixels := c.hDC.GetDeviceCaps(PHYSICALOFFSETY)
	xform := NewXFORM(float32(xDPI)/72, float32(yDPI)/72, float32(-xMarginPixels), float32(-yMarginPixels))
	if err := c.hDC.SetGraphicsMode(GM_ADVANCED); err != nil {
		return err
	}
	if err := c.hDC.SetWorldTransform(xform); err != nil {
		return err
	}

	if err := c.hDC.StartPage(); err != nil {
		return err
	}
	defer c.hDC.EndPage()

	if err := c.cContext.Save(); err != nil {
		return err
	}

	wPaperPixels := c.hDC.GetDeviceCaps(PHYSICALWIDTH)
	hPaperPixels := c.hDC.GetDeviceCaps(PHYSICALHEIGHT)
	wPrintablePixels := c.hDC.GetDeviceCaps(HORZRES)
	hPrintablePixels := c.hDC.GetDeviceCaps(VERTRES)

	wDocPoints, hDocPoints, err := pPage.GetSize()
	if err != nil {
		return err
	}

	scale, xOffsetPoints, yOffsetPoints := getScaleAndOffset(wDocPoints, hDocPoints, wPaperPixels, hPaperPixels, xMarginPixels, yMarginPixels, wPrintablePixels, hPrintablePixels, xDPI, yDPI, fitToPage)

	if err := c.cContext.IdentityMatrix(); err != nil {
		return err
	}
	if err := c.cContext.Translate(xOffsetPoints, yOffsetPoints); err != nil {
		return err
	}
	if err := c.cContext.Scale(scale, scale); err != nil {
		return err
	}

	pPage.RenderForPrinting(c.cContext)

	if err := c.cContext.Restore(); err != nil {
		return err
	}
	if err := c.cSurface.ShowPage(); err != nil {
		return err
	}

	return nil
}

var (
	colorValueByType = map[model.ColorType]int16{
		model.ColorTypeStandardColor:      DMCOLOR_COLOR,
		model.ColorTypeStandardMonochrome: DMCOLOR_MONOCHROME,
		// Ignore the rest, since we don't advertise them.
	}

	duplexValueByType = map[model.DuplexType]int16{
		model.DuplexNoDuplex:  DMDUP_SIMPLEX,
		model.DuplexLongEdge:  DMDUP_VERTICAL,
		model.DuplexShortEdge: DMDUP_HORIZONTAL,
	}

	pageOrientationByType = map[model.PageOrientationType]int16{
		model.PageOrientationPortrait:  DMORIENT_PORTRAIT,
		model.PageOrientationLandscape: DMORIENT_LANDSCAPE,
		// Ignore model.PageOrientationAuto for ticket parsing, in order to interpret "auto".
	}
)

// Print sends a new print job to the specified printer. The job ID
// is returned.
func (ws *WinSpool) Print(printer *lib.Printer, fileName, title string, ticket *model.JobTicket) (uint32, error) {
	printer.NativeJobSemaphore.Acquire()
	defer printer.NativeJobSemaphore.Release()

	if printer == nil {
		return 0, errors.New("Print() called with nil printer")
	}
	if ticket == nil {
		return 0, errors.New("Print() called with nil ticket")
	}

	jobContext, err := newJobContext(printer.Name, fileName, title)
	if err != nil {
		return 0, err
	}
	defer jobContext.free()

	if ticket.Color != nil && printer.Description.Color != nil {
		if color, ok := colorValueByType[ticket.Color.Type]; ok {
			jobContext.devMode.SetColor(color)
		} else if ticket.Color.VendorID != "" {
			v, err := strconv.ParseInt(ticket.Color.VendorID, 10, 16)
			if err != nil {
				return 0, err
			}
			jobContext.devMode.SetColor(int16(v))
		}
	}

	if ticket.Duplex != nil && printer.Description.Duplex != nil {
		if duplex, ok := duplexValueByType[ticket.Duplex.Type]; ok {
			jobContext.devMode.SetDuplex(duplex)
		}
	}

	if ticket.PageOrientation != nil && printer.Description.PageOrientation != nil {
		if pageOrientation, ok := pageOrientationByType[ticket.PageOrientation.Type]; ok {
			jobContext.devMode.SetOrientation(pageOrientation)
		}
	}

	if ticket.Copies != nil && printer.Description.Copies != nil {
		if ticket.Copies.Copies > 0 {
			jobContext.devMode.SetCopies(int16(ticket.Copies.Copies))
		}
	}

	var fitToPage bool
	if ticket.FitToPage != nil && printer.Description.FitToPage != nil {
		if ticket.FitToPage.Type == model.FitToPageFitToPage {
			fitToPage = true
		}
	}

	if ticket.MediaSize != nil && printer.Description.MediaSize != nil {
		if ticket.MediaSize.VendorID != "" {
			v, err := strconv.ParseInt(ticket.MediaSize.VendorID, 10, 16)
			if err != nil {
				return 0, err
			}
			jobContext.devMode.SetPaperSize(int16(v))
			jobContext.devMode.ClearPaperLength()
			jobContext.devMode.ClearPaperWidth()
		} else {
			jobContext.devMode.ClearPaperSize()
			jobContext.devMode.SetPaperLength(int16(ticket.MediaSize.HeightMicrons / 10))
			jobContext.devMode.SetPaperWidth(int16(ticket.MediaSize.WidthMicrons / 10))
		}
	}

	if ticket.Collate != nil && printer.Description.Collate != nil {
		if ticket.Collate.Collate {
			jobContext.devMode.SetCollate(DMCOLLATE_TRUE)
		} else {
			jobContext.devMode.SetCollate(DMCOLLATE_FALSE)
		}
	}

	for i := 0; i < jobContext.pDoc.GetNPages(); i++ {
		if err := printPage(printer.Name, i, jobContext, fitToPage); err != nil {
			return 0, err
		}
	}

	// Retain unpaused jobs to check the status later. Don't retain paused jobs because
	// release would delete the job even if it was still paused and hadn't been printed
	ji1, err := jobContext.hPrinter.GetJob(jobContext.jobID)
	if err != nil {
		return 0, err
	}
	if ji1.status&JOB_STATUS_PAUSED == 0 {
		err = jobContext.hPrinter.SetJobCommand(jobContext.jobID, JOB_CONTROL_RETAIN)
		if err != nil {
			return 0, err
		}
	}

	return uint32(jobContext.jobID), nil
}

func (ws *WinSpool) ReleaseJob(printerName string, jobID uint32) error {
	hPrinter, err := OpenPrinter(printerName)
	if err != nil {
		return err
	}

	// Only release if the job was retained (otherwise we get an error)
	ji1, err := hPrinter.GetJob(int32(jobID))
	if err != nil {
		return err
	}
	if ji1.status&JOB_STATUS_RETAINED != 0 {
		err = hPrinter.SetJobCommand(int32(jobID), JOB_CONTROL_RELEASE)
		if err != nil {
			return err
		}
	}

	return nil
}

type Job struct {
	Status         uint32
	Priority       uint32
	Size           uint32
	PrinterName    string
	DriverName     string
	Document       string
	PrintProcessor string
	Datatype       string
	JobID          uint32
	MachineName    string
	UserName       string
}

func (ws *WinSpool) JobList(printerName string) ([]Job, error) {
	hPrinter, err := OpenPrinter(printerName)
	if err != nil {
		return nil, err
	}
	jobs1, err := hPrinter.EnumJobs1()
	jobs := make([]Job, len(jobs1))
	for i := range jobs1 {
		jobs[i] = Job{
			Document:    utf16PtrToString(jobs1[i].pDocument),
			MachineName: utf16PtrToString(jobs1[i].pMachineName),
			Datatype:    utf16PtrToString(jobs1[i].pDatatype),
			PrinterName: utf16PtrToString(jobs1[i].pPrinterName),
			UserName:    utf16PtrToString(jobs1[i].pUserName),
			Status:      jobs1[i].status,
			Priority:    jobs1[i].priority,
			JobID:       jobs1[i].jobID,
		}

	}
	return jobs, err
}

func (ws *WinSpool) StartPrinterNotifications(handle windows.Handle) error {
	err := RegisterDeviceNotification(handle)
	return err
}

// The following functions are not relevant to Windows printing, but are required by the NativePrintSystem interface.

func (ws *WinSpool) RemoveCachedPPD(printerName string) {}
