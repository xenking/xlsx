package xlsx

import (
	"bytes"
	"errors"
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/xenking/redis"
)

type RedisRow struct {
	row         *Row
	maxCol      int
	client      *redis.Client
	buf         bytes.Buffer
	currentCell *Cell
}

func makeRedisRow(sheet *Sheet, client *redis.Client) *RedisRow {
	rr := &RedisRow{
		row:    new(Row),
		maxCol: -1,
		client: client,
	}
	rr.row.Sheet = sheet
	rr.row.cellStoreRow = rr
	sheet.setCurrentRow(rr.row)
	return rr
}

func (rr *RedisRow) CellUpdatable(c *Cell) {
	if c != rr.currentCell {
		panic("Attempt to update Cell that isn't the current cell whilst using the RedisCellStore.  You must use the Cell returned by the most recent operation.")

	}
}
func (rr *RedisRow) Updatable() {
	if rr.row != rr.row.Sheet.currentRow {
		panic("Attempt to update Row that isn't the current row whilst using the RedisCellStore.  You must use the row returned by the most recent operation.")
	}
}

func (rr *RedisRow) AddCell() *Cell {
	cell := newCell(rr.row, rr.maxCol+1)
	rr.setCurrentCell(cell)
	return cell
}

func (rr *RedisRow) readCell(index int) (*Cell, error) {
	var err error
	var cellType int
	var hasStyle, hasDataValidation bool
	var cellIsNil bool
	key := rr.row.makeCellKeyPrefix(index)
	b, err := rr.client.HGET(key, rr.row.makeRowNum())
	if err != nil {
		return nil, err
	}

	buf := bytes.NewReader(b)
	if cellIsNil, err = readBool(buf); err != nil {
		return nil, err
	}
	if cellIsNil {
		if err = readEndOfRecord(buf); err != nil {
			return nil, err
		}
		return nil, nil
	}
	c := &Cell{}
	if c.Value, err = readString(buf); err != nil {
		return c, err
	}
	if c.formula, err = readString(buf); err != nil {
		return c, err
	}
	if hasStyle, err = readBool(buf); err != nil {
		return c, err
	}
	if c.NumFmt, err = readString(buf); err != nil {
		return c, err
	}
	if c.date1904, err = readBool(buf); err != nil {
		return c, err
	}
	if c.Hidden, err = readBool(buf); err != nil {
		return c, err
	}
	if c.HMerge, err = readInt(buf); err != nil {
		return c, err
	}
	if c.VMerge, err = readInt(buf); err != nil {
		return c, err
	}
	if cellType, err = readInt(buf); err != nil {
		return c, err
	}
	c.cellType = CellType(cellType)
	if hasDataValidation, err = readBool(buf); err != nil {
		return c, err
	}
	if c.Hyperlink.DisplayString, err = readString(buf); err != nil {
		return c, err
	}
	if c.Hyperlink.Link, err = readString(buf); err != nil {
		return c, err
	}
	if c.Hyperlink.Tooltip, err = readString(buf); err != nil {
		return c, err
	}
	if c.num, err = readInt(buf); err != nil {
		return c, err
	}
	if c.RichText, err = readRichText(buf); err != nil {
		return c, err
	}
	if err = readEndOfRecord(buf); err != nil {
		return c, err
	}
	if hasStyle {
		if c.style, err = readStyle(buf); err != nil {
			return c, err
		}
	}
	if hasDataValidation {
		if c.DataValidation, err = readDataValidation(buf); err != nil {
			return c, err
		}
	}
	return c, nil
}

func (rr *RedisRow) writeCell(c *Cell) error {
	var err error
	rr.buf.Reset()
	if c == nil {
		if err := writeBool(&rr.buf, true); err != nil {

			return err
		}
		return writeEndOfRecord(&rr.buf)
	}
	if err := writeBool(&rr.buf, false); err != nil {
		return err
	}
	if err = writeString(&rr.buf, c.Value); err != nil {
		return err
	}
	if err = writeString(&rr.buf, c.formula); err != nil {
		return err
	}
	if err = writeBool(&rr.buf, c.style != nil); err != nil {
		return err
	}
	if err = writeString(&rr.buf, c.NumFmt); err != nil {
		return err
	}
	if err = writeBool(&rr.buf, c.date1904); err != nil {
		return err
	}
	if err = writeBool(&rr.buf, c.Hidden); err != nil {
		return err
	}
	if err = writeInt(&rr.buf, c.HMerge); err != nil {
		return err
	}
	if err = writeInt(&rr.buf, c.VMerge); err != nil {
		return err
	}
	if err = writeInt(&rr.buf, int(c.cellType)); err != nil {
		return err
	}
	if err = writeBool(&rr.buf, c.DataValidation != nil); err != nil {
		return err
	}
	if err = writeString(&rr.buf, c.Hyperlink.DisplayString); err != nil {
		return err
	}
	if err = writeString(&rr.buf, c.Hyperlink.Link); err != nil {
		return err
	}
	if err = writeString(&rr.buf, c.Hyperlink.Tooltip); err != nil {
		return err
	}
	if err = writeInt(&rr.buf, c.num); err != nil {
		return err
	}
	if err = writeRichText(&rr.buf, c.RichText); err != nil {
		return err
	}
	if err = writeEndOfRecord(&rr.buf); err != nil {
		return err
	}
	if c.style != nil {
		if err = writeStyle(&rr.buf, c.style); err != nil {
			return err
		}
	}
	if c.DataValidation != nil {
		if err = writeDataValidation(&rr.buf, c.DataValidation); err != nil {
			return err
		}
	}
	key := rr.row.makeCellKeyPrefix(c.num)
	_, err = rr.client.ZADDString(makeSheetCellsStore(rr.row.Sheet.Name), int64(c.num), key)
	if err != nil {
		return err
	}
	_, err = rr.client.HSET(key, rr.row.makeRowNum(), rr.buf.Bytes())
	return err
}

func (rr *RedisRow) setCurrentCell(cell *Cell) {
	if rr.currentCell.Modified() {
		err := rr.writeCell(rr.currentCell)
		if err != nil {
			panic(err.Error())
		}
	}
	if cell.num > rr.maxCol {
		rr.maxCol = cell.num
	}
	rr.currentCell = cell

}

func (rr *RedisRow) PushCell(c *Cell) {
	c.modified = true
	rr.setCurrentCell(c)
}

func (rr *RedisRow) GetCell(colIdx int) *Cell {
	if rr.currentCell != nil {
		if rr.currentCell.num == colIdx {
			return rr.currentCell
		}
	}
	cell, err := rr.readCell(colIdx)
	if err == nil {
		rr.setCurrentCell(cell)
		return cell
	}
	cell = newCell(rr.row, colIdx)
	rr.PushCell(cell)
	return cell
}

func (rr *RedisRow) ForEachCell(cvf CellVisitorFunc, option ...CellVisitorOption) error {
	flags := &cellVisitorFlags{}
	for _, opt := range option {
		opt(flags)
	}
	fn := func(ci int, c *Cell) error {
		if c == nil {
			if flags.skipEmptyCells {
				return nil
			}
			c = rr.GetCell(ci)
		}
		if !c.Modified() && flags.skipEmptyCells {
			return nil
		}
		c.Row = rr.row
		rr.setCurrentCell(c)
		return cvf(c)
	}

	for ci := 0; ci <= rr.maxCol; ci++ {
		var cell *Cell
		key := rr.row.makeCellKeyPrefix(ci)
		b, err := rr.client.HGET(key, rr.row.makeRowNum())
		if err != nil {
			// If the file doesn't exist that's fine, it was just an empty cell.
			if !os.IsNotExist(err) {
				return err
			}

		} else {
			cell, err = readCell(bytes.NewReader(b))
			if err != nil {
				return err
			}
		}

		err = fn(ci, cell)
		if err != nil {
			return err
		}
	}

	if !flags.skipEmptyCells {
		for ci := rr.maxCol + 1; ci < rr.row.Sheet.MaxCol; ci++ {
			c := rr.GetCell(ci)
			err := cvf(c)
			if err != nil {
				return err
			}
		}
	}

	return nil
}

// MaxCol returns the index of the rightmost cell in the row's column.
func (rr *RedisRow) MaxCol() int {
	return rr.maxCol
}

// CellCount returns the total number of cells in the row.
func (rr *RedisRow) CellCount() int {
	return rr.maxCol + 1
}

// RedisCellStore is an implementation of the CellStore interface, backed by Redis
type RedisCellStore struct {
	sheetName string
	buf       *bytes.Buffer
	reader    *bytes.Reader
	client    *redis.Client
}

// UseRedisCellStore is a FileOption that makes all Sheet instances
// for a File use Redis as their backing client.  You can use this
// option when handling very large Sheets that would otherwise require
// allocating vast amounts of memory.
func UseRedisCellStore(options ...RedisCellStoreOption) FileOption {
	return func(f *File) {
		f.cellStoreConstructor = NewRedisCellStoreConstructor(options...)
	}
}

type RedisCellStoreOption struct {
	RedisAddr string
	CommandTimeout time.Duration
	DialTimeout time.Duration
}

// NewRedisCellStoreConstructor is a CellStoreConstructor than returns a
// CellStore in terms of Redis.
func NewRedisCellStoreConstructor(options RedisCellStoreOption) CellStoreConstructor {
	return func() (CellStore, error) {
		cs := &RedisCellStore{
			buf: bytes.NewBuffer([]byte{}),
		}
		cs.client = redis.NewClient(options.RedisAddr, options.CommandTimeout, options.DialTimeout)
		return cs, nil
	}
}

// ReadRow reads a row from the persistent client, identified by key,
// into memory and returns it, with the provided Sheet set as the Row's Sheet.
func (cs *RedisCellStore) ReadRow(key string, s *Sheet) (*Row, error) {
	if len(cs.sheetName) == 0 && s != nil {
		cs.sheetName = s.Name
	}
	str := strings.Split(key, ":")
	if len(str) != 2 {
		return nil, NewRowNotFoundError(key, "no such row")
	}
	b, err := cs.client.HGET(makeSheetRowsStore(s.Name), str[1])
	if err != nil {
		return nil, err
	}
	if b == nil {
		return nil, NewRowNotFoundError(key, "no such row")
	}
	r, maxCol, err := readRedisRow(bytes.NewReader(b))
	if err != nil {
		return nil, err
	}
	r.Sheet = s
	dr := &RedisRow{
		row:    r,
		maxCol: maxCol,
		client: cs.client,
	}
	r.cellStoreRow = dr
	return r, nil
}

// MoveRow moves a Row from one position in a Sheet (index) to another
// within the persistent client.
func (cs *RedisCellStore) MoveRow(r *Row, index int) error {
	if len(cs.sheetName) == 0 && r.Sheet != nil {
		cs.sheetName = r.Sheet.Name
	}
	cell := r.cellStoreRow.(*RedisRow).currentCell
	if cell != nil {
		cs.buf.Reset()
		if err := writeCell(cs.buf, cell); err != nil {
			return err
		}
		key := r.makeCellKeyPrefix(cell.num)
		_, err := cs.client.ZADDString(makeSheetCellsStore(r.Sheet.Name), int64(cell.num), key)
		if err != nil {
			return err
		}
		if _, err := cs.client.HSET(key, r.makeRowNum(), cs.buf.Bytes()); err != nil {
			return err
		}
	}
	oldKey := r.makeRowNum()
	newKey := strconv.Itoa(index)
	val, err := cs.client.HGET(makeSheetRowsStore(r.Sheet.Name), newKey)
	if err != nil {
		return err
	}
	if val != nil {
		return fmt.Errorf("Target index for row (%d) would overwrite a row already exists", index)
	}
	_, err = cs.client.HDEL(makeSheetRowsStore(r.Sheet.Name), oldKey)
	if err != nil {
		return err
	}
	cs.buf.Reset()
	var cBuf bytes.Buffer
	err = r.ForEachCell(func(c *Cell) error {
		cBuf.Reset()
		k := r.makeCellKeyPrefix(c.num)
		c.Row = r
		err = writeCell(&cBuf, c)
		_, err = cs.client.HSET(k, newKey, cBuf.Bytes())
		if err != nil {
			return err
		}
		_, err = cs.client.HDEL(k, oldKey)
		return err
	}, SkipEmptyCells)
	if err != nil {
		return err
	}
	r.num = index
	err = writeRow(cs.buf, r)
	if err != nil {
		return err
	}
	_, err = cs.client.HSET(makeSheetRowsStore(r.Sheet.Name), newKey, cs.buf.Bytes())
	return err
}

// RemoveRow removes a Row from the Sheet's representation in the
// persistent client.
func (cs *RedisCellStore) RemoveRow(key string) error {
	k := strings.Split(key, ":")
	if len(k) != 2 {
		return NewRowNotFoundError(key, "no such row")
	}
	cells, err := cs.client.ZRANGEString(makeSheetCellsStore(k[0]), 0, -1)
	for _, cell := range cells {
		_, err = cs.client.HDEL(cell, k[1])
		if err != nil {
			return err
		}
	}
	_, err = cs.client.HDEL(k[0], k[1])
	if err != nil {
		return err
	}

	return nil
}

// MakeRow returns an empty Row
func (cs *RedisCellStore) MakeRow(sheet *Sheet) *Row {
	if len(cs.sheetName) == 0 && sheet != nil {
		cs.sheetName = sheet.Name
	}
	return makeRedisRow(sheet, cs.client).row
}

// MakeRowWithLen returns an empty Row, with a preconfigured starting length.
func (cs *RedisCellStore) MakeRowWithLen(sheet *Sheet, len int) *Row {
	mr := makeRedisRow(sheet, cs.client)
	mr.maxCol = len - 1
	return mr.row
}

func readRedisRow(reader *bytes.Reader) (*Row, int, error) {
	var err error
	var maxCol int

	r := &Row{}

	r.Hidden, err = readBool(reader)
	if err != nil {
		return nil, maxCol, err
	}
	height, err := readFloat(reader)
	if err != nil {
		return nil, maxCol, err
	}
	r.height = height
	outlineLevel, err := readInt(reader)
	if err != nil {
		return nil, maxCol, err
	}
	r.outlineLevel = uint8(outlineLevel)
	r.isCustom, err = readBool(reader)
	if err != nil {
		return nil, maxCol, err
	}
	r.num, err = readInt(reader)
	if err != nil {
		return nil, maxCol, err
	}
	maxCol, err = readInt(reader)
	if err != nil {
		return nil, maxCol, err
	}
	err = readEndOfRecord(reader)
	if err != nil {
		return r, maxCol, err
	}
	return r, maxCol, nil
}

// Close will remove the persisant storage for a given Sheet completely.
func (cs *RedisCellStore) Close() error {
	cells, err := cs.client.ZRANGEString(makeSheetCellsStore(cs.sheetName), 0, -1)
	if err != nil {
		return err
	}
	_, err = cs.client.DELArgs(cells...)
	if err != nil {
		return err
	}
	_, err = cs.client.DEL(makeSheetRowsStore(cs.sheetName))
	if err != nil {
		return err
	}
	_, err = cs.client.DEL(makeSheetCellsStore(cs.sheetName))
	if err != nil {
		return err
	}
	return cs.client.Close()
}

// WriteRow writes a Row to persistant storage.
func (cs *RedisCellStore) WriteRow(r *Row) error {
	if len(cs.sheetName) == 0 && r.Sheet != nil {
		cs.sheetName = r.Sheet.Name
	}
	rr, ok := r.cellStoreRow.(*RedisRow)
	if !ok {
		return fmt.Errorf("cellStoreRow for a RedisCellStore is not RedisRow (%T)", r.cellStoreRow)
	}
	if rr.currentCell != nil {
		err := rr.writeCell(rr.currentCell)
		if err != nil {
			return err
		}
	}
	cs.buf.Reset()
	err := writeRow(cs.buf, r)
	if err != nil {
		return err
	}
	_, err = cs.client.HSET(makeSheetRowsStore(r.Sheet.Name), r.makeRowNum(), cs.buf.Bytes())
	return err
}
