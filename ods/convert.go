package ods

import (
	"errors"
	"image/color"
	"strconv"
	"strings"
)

var errInvalidFormat = errors.New("invalid format")

// ParseHexColor is from https://stackoverflow.com/questions/54197913/parse-hex-string-to-image-color
func ParseHexColor(s string) (c color.RGBA, err error) {
	c.A = 0xff

	if len(s) < 4 {
		return color.RGBA{}, nil
	}

	if s[0] != '#' {
		return c, errInvalidFormat
	}

	hexToByte := func(b byte) byte {
		switch {
		case b >= '0' && b <= '9':
			return b - '0'
		case b >= 'a' && b <= 'f':
			return b - 'a' + 10
		case b >= 'A' && b <= 'F':
			return b - 'A' + 10
		}
		err = errInvalidFormat
		return 0
	}

	switch len(s) {
	case 7:
		c.R = hexToByte(s[1])<<4 + hexToByte(s[2])
		c.G = hexToByte(s[3])<<4 + hexToByte(s[4])
		c.B = hexToByte(s[5])<<4 + hexToByte(s[6])
	case 4:
		c.R = hexToByte(s[1]) * 17
		c.G = hexToByte(s[2]) * 17
		c.B = hexToByte(s[3]) * 17
	default:
		err = errInvalidFormat
	}
	return
}

// ToMM function assumes, that measurements are in cm.
// TODO: is this always true? No, there may be pts.
// 1pt = 0.3527777778
func ToMM(s string) (float64, error) {
	if s == "" {
		return 0, nil
	}
	switch s[len(s)-2:] {
	case "cm":
		f, err := strconv.ParseFloat(strings.TrimSuffix(s, "cm"), 64)
		if err != nil {
			return 0, errors.New("ToMM: error converting string(cm) to float64, " + err.Error())
		}
		return f * 10, nil
	case "pt":
		f, err := strconv.ParseFloat(strings.TrimSuffix(s, "pt"), 64)
		if err != nil {
			return 0, errors.New("ToMM: error converting string(pt) to float64, " + err.Error())
		}
		return f * 0.3527777778, nil
	}
	return 0, errors.New("ToMM: unknown units! " + s)
}

func PxToFloat64(s string) (float64, error) {
	if s == "" {
		return 0, nil
	}
	f, err := strconv.ParseFloat(strings.TrimSuffix(s, "pt"), 64)
	if err != nil {
		return 0, errors.New("PxToFloat64: error converting string to float64," + err.Error())
	}
	return f, nil
}
