package xlref

import (
	"fmt"
	"strconv"
	"strings"
)

type RangeRef struct {
	Sheet     string
	StartCell string
	EndCell   string
}

func ParseA1Range(formula string) (RangeRef, error) {
	trimmed := strings.TrimSpace(formula)
	if trimmed == "" {
		return RangeRef{}, fmt.Errorf("empty formula")
	}
	if strings.HasPrefix(trimmed, "=") {
		trimmed = strings.TrimSpace(trimmed[1:])
	}

	sheet, cellPart, err := splitSheetAndCells(trimmed)
	if err != nil {
		return RangeRef{}, err
	}

	start, end, err := parseCellRange(cellPart)
	if err != nil {
		return RangeRef{}, err
	}

	return RangeRef{
		Sheet:     sheet,
		StartCell: start,
		EndCell:   end,
	}, nil
}

func splitSheetAndCells(formula string) (string, string, error) {
	if strings.HasPrefix(formula, "'") {
		sheet, rest, err := parseQuotedSheet(formula)
		if err != nil {
			return "", "", err
		}
		return sheet, rest, nil
	}

	parts := strings.SplitN(formula, "!", 2)
	if len(parts) != 2 {
		return "", "", fmt.Errorf("missing sheet separator")
	}

	sheet := strings.TrimSpace(parts[0])
	if sheet == "" {
		return "", "", fmt.Errorf("missing sheet name")
	}

	cellPart := strings.TrimSpace(parts[1])
	if cellPart == "" {
		return "", "", fmt.Errorf("missing cell reference")
	}

	return sheet, cellPart, nil
}

func parseQuotedSheet(formula string) (string, string, error) {
	sheet, rest, err := readQuotedSheet(formula)
	if err != nil {
		return "", "", err
	}

	rest = strings.TrimSpace(rest)
	if rest == "" || rest[0] != '!' {
		return "", "", fmt.Errorf("missing sheet separator")
	}

	cellPart := strings.TrimSpace(rest[1:])
	if cellPart == "" {
		return "", "", fmt.Errorf("missing cell reference")
	}

	return sheet, cellPart, nil
}

func readQuotedSheet(formula string) (string, string, error) {
	if formula == "" || formula[0] != '\'' {
		return "", "", fmt.Errorf("missing quoted sheet name")
	}

	var builder strings.Builder
	escaped := false
	for i := 1; i < len(formula); i++ {
		ch := formula[i]
		if ch == '\'' {
			if i+1 < len(formula) && formula[i+1] == '\'' {
				builder.WriteByte('\'')
				i++
				continue
			}
			escaped = true
			return builder.String(), formula[i+1:], nil
		}
		builder.WriteByte(ch)
	}

	if !escaped {
		return "", "", fmt.Errorf("unterminated sheet name")
	}
	return "", "", fmt.Errorf("unterminated sheet name")
}

func parseCellRange(cellPart string) (string, string, error) {
	parts := strings.Split(cellPart, ":")
	if len(parts) == 0 || len(parts) > 2 {
		return "", "", fmt.Errorf("invalid cell range")
	}

	start, err := parseCell(parts[0])
	if err != nil {
		return "", "", err
	}

	if len(parts) == 1 {
		return start, start, nil
	}

	end, err := parseCell(parts[1])
	if err != nil {
		return "", "", err
	}

	return start, end, nil
}

func parseCell(cell string) (string, error) {
	trimmed := strings.TrimSpace(cell)
	if trimmed == "" {
		return "", fmt.Errorf("empty cell reference")
	}

	trimmed = strings.ReplaceAll(trimmed, "$", "")
	if trimmed == "" {
		return "", fmt.Errorf("empty cell reference")
	}

	i := 0
	for i < len(trimmed) {
		ch := trimmed[i]
		if ch >= 'A' && ch <= 'Z' || ch >= 'a' && ch <= 'z' {
			i++
			continue
		}
		break
	}
	if i == 0 || i == len(trimmed) {
		return "", fmt.Errorf("invalid cell reference")
	}

	col := strings.ToUpper(trimmed[:i])
	rowStr := trimmed[i:]
	for _, ch := range rowStr {
		if ch < '0' || ch > '9' {
			return "", fmt.Errorf("invalid cell reference")
		}
	}

	row, err := strconv.Atoi(rowStr)
	if err != nil || row <= 0 {
		return "", fmt.Errorf("invalid cell reference")
	}

	return fmt.Sprintf("%s%d", col, row), nil
}
